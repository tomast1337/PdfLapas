import {
    assertHtmlPresentationDocument,
    assertLinkUrl,
    checkImagePixels,
    type HtmlPresentationDocument,
    type ImagePixelLimits,
    imagePixelLimits,
    inchesToPixels,
    type MissingFont,
    type MissingFontPolicy,
    missingFontPolicy,
    type PresentationRenderer,
    reportMissingFonts,
} from "@pdflapas/model";
import { withBrowser } from "./browser";
import { DEFAULT_LAUNCH_OPTIONS, type PuppeteerLaunchOptions } from "./launchOptions";
import { type DocumentInfo, withDocumentInfo } from "./pdfInfo";
import { LINK_PLACEHOLDER, withLinkUris } from "./pdfLinks";
import { metricSubstitute, PRINT_ZOOM, printHtml } from "./printWorkarounds";
import { restrictedFonts } from "./fontEmbedding";
import { type3FontNames, withType1CFonts } from "./type3Fonts";

/** Options for `PuppeteerPdfRenderer`. */
export type PuppeteerPdfRendererOptions = {
    /** Merged over `DEFAULT_LAUNCH_OPTIONS` (shallowly: `args` given here replaces the default args). */
    readonly launchOptions?: PuppeteerLaunchOptions;
    /**
     * What to do when a text run asks for a face Chromium has no font for (its metric-compatible fontconfig alias
     * counts as the face): `"warn"` (default), `"throw"` (`MissingFontError`), `"ignore"`, or a function receiving the list.
     */
    readonly onMissingFont?: MissingFontPolicy;
    /** Decoded-pixel limits on the deck's raster images, checked from their headers before printing. Default IMAGE_PIXEL_LIMITS. */
    readonly imageLimits?: ImagePixelLimits;
};

export class PuppeteerPdfRenderer implements PresentationRenderer {
    private readonly launchOptions: PuppeteerLaunchOptions;
    private readonly onMissingFont: MissingFontPolicy;
    private readonly imageLimits: Required<ImagePixelLimits>;

    constructor({ launchOptions, onMissingFont, imageLimits }: PuppeteerPdfRendererOptions = {}) {
        this.launchOptions = { ...DEFAULT_LAUNCH_OPTIONS, ...launchOptions };
        this.onMissingFont = missingFontPolicy(onMissingFont);
        this.imageLimits = imagePixelLimits(imageLimits);
    }

    async render(document: HtmlPresentationDocument): Promise<Uint8Array> {
        assertHtmlPresentationDocument(document);
        checkImagePixels(document.html, this.imageLimits);

        return withBrowser(async browser => {
            // A browser context per document: no cache, storage or cookies shared with another render.
            const context = await browser.createBrowserContext();

            try {
                const page = await context.newPage();

                // Cut the print page off from the network and turn off its scripts, so markup that names a URL (a CSS
                // url(), an iframe, a link, a script) reaches no host from the server. page.evaluate runs through the
                // DevTools protocol, which these settings do not stop, so the print steps below still run.
                await page.setJavaScriptEnabled(false);
                await page.setRequestInterception(true);
                page.on("request", request => {
                    if (ALLOWED_REQUEST.test(request.url())) {
                        void request.continue();
                    } else {
                        void request.abort("blockedbyclient");
                    }
                });

                const { html, pageSize } = document;

                await page.setViewport({
                    width: Math.round(inchesToPixels(pageSize.width)),
                    height: Math.round(inchesToPixels(pageSize.height)),
                });
                // Images are inlined by resolveDocumentImages; prepareForPrint waits for fonts and decoding.
                // The print workarounds are in the document before anything measures it: its text shrinking scales them.
                await page.setContent(printHtml(html), { waitUntil: "load" });
                const failed = await page.evaluate(prepareForPrint);

                if (failed.length) {
                    throw new Error(`Image failed to decode: ${failed.join(", ")}`);
                }

                // Before printing, so "throw" does not print first; the fonts are loaded once prepareForPrint returns. A Calibri or
                // Cambria face drawn in its metric-compatible substitute is drawn as intended (the theme's Calibri Light in Carlito).
                const missing = await page.evaluate(missingFonts);

                reportMissingFonts(
                    missing.filter(({ face, usedFace }) => {
                        const substitute = metricSubstitute(face);

                        return !usedFace || substitute.kind === "none" || usedFace !== substitute.face;
                    }),
                    this.onMissingFont,
                );
                const links = await page.evaluate(placeLinkPlaceholders, LINK_PLACEHOLDER);

                for (const { href, object } of links) {
                    assertLinkUrl(href, object);
                }

                const uris = links.map(({ href }) => href);

                const pdf = await page.pdf({
                    width: `${pageSize.width}in`,
                    height: `${pageSize.height}in`,
                    // The document's @page size wins over width and height, which are only the fallback. Chromium rounds a
                    // width/height paper up and centres the @page box on it at a whole CSS px, shifting everything by up to
                    // 1px (0.75pt) on sizes such as 8.2646 x 11.695in; LibreOffice and PowerPoint place it at the origin.
                    preferCSSPageSize: true,
                    // printHtml lays the document out at PRINT_ZOOM; this scales it back to the deck size.
                    scale: 1 / PRINT_ZOOM,
                    printBackground: true,
                    margin: { top: 0, right: 0, bottom: 0, left: 0 },
                });
                // The links update reads the newest cross-reference section, so it follows the Info update, which needs the original
                // one; the fonts update follows every section.
                const linked = withLinkUris(withDocumentInfo(pdf, await page.evaluate(documentInfo)), uris);

                return withType1CFonts(linked, await restrictedFonts(type3FontNames(linked)));
            } finally {
                await context.close();
            }
        }, this.launchOptions);
    }
}

// The print page loads only inline content: resolveDocumentImages has turned every image into a data: URI.
const ALLOWED_REQUEST = /^(?:data:|about:blank$)/;

// prepareForPrint and documentInfo run inside the page: page.evaluate sends their source text, so they may use only
// their own code (prepareForPrint's stages are written in it for that reason) and the page's globals.

/**
 * Settles the document for printing in one page.evaluate round trip, and returns the names of the images that did
 * not decode.
 */
async function prepareForPrint(): Promise<string[]> {
    await document.fonts.ready;
    await Promise.all(Array.from(document.images, settled));

    // LibreOffice's PDF export drops animated GIFs, so it cannot be used as the visual oracle here. Freeze through
    // Chromium's static bitmap decoder to match PowerPoint's PDF first frame.
    const gifs = Array.from(document.images).filter(image => image.src.startsWith("data:image/gif"));

    if (gifs.length) {
        if (typeof createImageBitmap !== "function") {
            throw new Error("This Chromium build cannot freeze animated GIF images");
        }

        await Promise.all(
            gifs.map(async image => {
                const frame = await createImageBitmap(await (await fetch(image.src)).blob());
                const canvas = document.createElement("canvas");

                canvas.width = frame.width;
                canvas.height = frame.height;
                canvas.getContext("2d")?.drawImage(frame, 0, 0);
                image.src = canvas.toDataURL("image/png");
                frame.close();
            }),
        );
    }

    // The frozen GIFs decode again from their new PNG source.
    await Promise.all(gifs.map(settled));

    // PowerPoint scales the decoded image surface into its declared blip rectangle. Explicitly transform from
    // natural pixels so SVG preserveAspectRatio rules cannot override that geometry.
    for (const image of Array.from(document.querySelectorAll<HTMLImageElement>(".slide-image-content"))) {
        if (!image.naturalWidth || !image.naturalHeight) {
            continue;
        }

        const computed = getComputedStyle(image);
        const targetWidth = parseFloat(computed.width);
        const targetHeight = parseFloat(computed.height);

        if (!targetWidth || !targetHeight) {
            continue;
        }

        image.style.width = `${image.naturalWidth}px`;
        image.style.height = `${image.naturalHeight}px`;
        image.style.transformOrigin = "top left";
        image.style.transform = `scale(${targetWidth / image.naturalWidth}, ${targetHeight / image.naturalHeight})`;
    }

    // The horizontal paragraphs whose last line (data-mark-baseline) or lines ending in a soft break
    // (data-break-baseline) sit higher in their line box than data-line-baseline, where the print step put every line,
    // with those raises as fractions of the line height: read before the text shrinking below scales the line height.
    const marked: { paragraph: HTMLElement; mark: number; soft: number }[] = [];

    for (const paragraph of Array.from(document.querySelectorAll<HTMLElement>(".text-paragraph[data-mark-baseline], .text-paragraph[data-break-baseline]"))) {
        const content = paragraph.parentElement;

        if (!content?.classList.contains("text-content") || content.parentElement?.dataset.textDirection !== "horizontal") {
            continue;
        }

        const { lineBaseline = "", markBaseline = lineBaseline, breakBaseline = lineBaseline } = paragraph.dataset;
        const [baseline, mark, soft, height] = [parseFloat(lineBaseline), parseFloat(markBaseline), parseFloat(breakBaseline), parseFloat(paragraph.style.lineHeight)] as const;

        if (!Number.isFinite(baseline) || !Number.isFinite(mark) || !Number.isFinite(soft) || !height) {
            const name = paragraph.closest<HTMLElement>("[data-object-name]")?.dataset.objectName ?? "unnamed text box";

            throw new Error(`Cannot print '${name}': a .text-paragraph with data-mark-baseline or data-break-baseline needs numeric data-line-baseline, data-mark-baseline and data-break-baseline (px) and a px line-height (PdfLapas HTML profile)`);
        }

        marked.push({ paragraph, mark: (baseline - mark) / height, soft: (baseline - soft) / height });
    }

    // Shrinks the text of `fit: "shrink"` boxes until it fits inside the box padding, by the rule the box states
    // (docs/HTML_PROFILE.md, Text): run and bullet font sizes, px line heights and the baseline shift scale together.
    for (const box of Array.from(document.querySelectorAll<HTMLElement>('.slide-text[data-text-fit="shrink"]'))) {
        const content = box.querySelector<HTMLElement>(".text-content");

        if (!content) {
            continue;
        }

        const step = Number(box.dataset.textFitStep);
        const maxSteps = Number(box.dataset.textFitMaxSteps);
        const tolerance = Number(box.dataset.textFitTolerance);

        if (!(step > 0 && step < 1) || !Number.isInteger(maxSteps) || maxSteps < 0 || !(tolerance >= 0)) {
            const name = box.dataset.objectName ?? "unnamed text box";

            throw new Error(
                `Cannot print '${name}': a data-text-fit="shrink" box needs data-text-fit-step (between 0 and 1), data-text-fit-max-steps (a whole number) and data-text-fit-tolerance (px); complete hand-written HTML with completeProfileHtml()`,
            );
        }

        const boxStyle = getComputedStyle(box);
        const availableWidth = box.clientWidth - parseFloat(boxStyle.paddingLeft) - parseFloat(boxStyle.paddingRight);
        const availableHeight = box.clientHeight - parseFloat(boxStyle.paddingTop) - parseFloat(boxStyle.paddingBottom);

        for (let iteration = 0; iteration < maxSteps && (content.scrollWidth > availableWidth + tolerance || content.scrollHeight > availableHeight + tolerance); iteration++) {
            for (const run of Array.from(content.querySelectorAll<HTMLElement>(".text-run, .text-bullet"))) {
                run.style.fontSize = `${parseFloat(getComputedStyle(run).fontSize) * step}px`;
            }

            for (const paragraph of Array.from(content.querySelectorAll<HTMLElement>(".text-paragraph"))) {
                const lineHeight = paragraph.style.lineHeight;

                if (lineHeight.endsWith("px")) {
                    paragraph.style.lineHeight = `${parseFloat(lineHeight) * step}px`;
                }

                // The baseline shift is a fraction of the line box and font size, so it shrinks with them.
                // shortcut: the scaled shift keeps the unscaled line's strut floor (printWorkarounds strutBaseline),
                // up to 1 device px off; recompute it from data-line-baseline when shrunk baselines are measured.
                if (paragraph.style.top.endsWith("px")) {
                    paragraph.style.top = `${parseFloat(paragraph.style.top) * step}px`;
                }
            }
        }
    }

    // Raises the glyphs of each marked paragraph's last line by its mark raise and of each line ending in a soft break by
    // its break raise. Which characters are on that line is known only after layout, so the text from the line's first
    // character on is wrapped in relatively positioned spans, which move the glyphs and leave every line break as it is.
    {
        // Client rects are in zoomed px (the print zoom on <html>), line heights in CSS px.
        const zoom = parseFloat(getComputedStyle(document.documentElement).zoom) || 1;
        const range = document.createRange();
        const moves: { nodes: Text[]; offset: number; raise: number; bullet?: HTMLElement }[] = [];
        /**
         * The vertical centre of the character at `offset` in `node`, in client px, where its line puts it: without the
         * `top` that raises or lowers its run (#298), which moves the glyph and not the line. None when it draws no box.
         */
        const centre = (node: Text, offset: number): { kind: "centre"; y: number } | { kind: "none" } => {
            range.setStart(node, offset);
            range.setEnd(node, offset + 1);
            const rect = range.getClientRects()[0];
            const shift = node.parentElement ? parseFloat(getComputedStyle(node.parentElement).top) || 0 : 0;

            return rect ? { kind: "centre", y: rect.top + rect.height / 2 - shift * zoom } : { kind: "none" };
        };

        // All reads first, so the wrapping below does not lay the document out again for each paragraph.
        for (const { paragraph, mark, soft } of marked) {
            const height = parseFloat(paragraph.style.lineHeight);
            const segments: Text[][] = [[]];

            for (const child of Array.from(paragraph.children)) {
                if (child.tagName === "BR") {
                    segments.push([]);
                } else if (child.classList.contains("text-run")) {
                    // Its own text and its words' spans' (wordCharacterSpacing), inside the print step's strike span or not, in order.
                    const segment = segments[segments.length - 1];
                    const walker = document.createTreeWalker(child, NodeFilter.SHOW_TEXT);

                    for (let text = walker.nextNode(); text; text = walker.nextNode()) {
                        if (text instanceof Text && text.length > 0) {
                            segment?.push(text);
                        }
                    }
                }
            }

            for (const [index, nodes] of segments.entries()) {
                const raise = (index === segments.length - 1 ? mark : soft) * height;

                if (!raise) {
                    continue;
                }

                // Where the line of the last character of `nodes` starts: the node index and offset of its first character.
                // Each line is one line height tall, so a character is on that line when its centre is within half a line
                // height of the last one's.
                const band = (height * zoom) / 2;
                const lastNode = nodes[nodes.length - 1];

                if (!lastNode) {
                    continue;
                }

                const last = centre(lastNode, lastNode.length - 1);

                if (last.kind === "none") {
                    continue;
                }

                const onLine = (node: Text, offset: number) => {
                    const y = centre(node, offset);

                    return y.kind === "none" || Math.abs(y.y - last.y) < band;
                };
                let start = { index: nodes.length - 1, offset: lastNode.length - 1 };

                for (let nodeIndex = nodes.length - 1; nodeIndex >= 0; nodeIndex--) {
                    const node = nodes[nodeIndex];

                    if (!node || !onLine(node, node.length - 1)) {
                        break;
                    }

                    if (onLine(node, 0)) {
                        start = { index: nodeIndex, offset: 0 };
                        continue;
                    }

                    // Lines only move down through the text: offset `low` is on an earlier line, `high` on this one.
                    let [low, high] = [0, node.length - 1];

                    while (high - low > 1) {
                        const middle = Math.floor((low + high) / 2);

                        if (onLine(node, middle)) {
                            high = middle;
                        } else {
                            low = middle;
                        }
                    }

                    start = { index: nodeIndex, offset: high };
                    break;
                }

                const onFirstLine = index === 0 && start.index === 0 && start.offset === 0;
                const bullet = onFirstLine ? (paragraph.querySelector<HTMLElement>(":scope > .text-bullet") ?? undefined) : undefined;

                moves.push({ nodes: nodes.slice(start.index), offset: start.offset, raise, bullet });
            }
        }

        for (const { nodes, offset, raise, bullet } of moves) {
            for (const [index, node] of nodes.entries()) {
                const span = document.createElement("span");

                span.style.position = "relative";
                span.style.top = `${-raise}px`;
                const moved = index === 0 && offset > 0 ? node.splitText(offset) : node;

                moved.replaceWith(span);
                span.append(moved);
            }

            if (bullet) {
                bullet.style.position = "relative";
                bullet.style.top = `${-raise}px`;
            }
        }
    }

    // Justifies each `text-justify: inter-word` paragraph (docs/HTML_PROFILE.md, Text) at its spaces only, which Chromium
    // has no `text-justify` for: its `auto` widens every Han character too. The paragraph is laid out left-aligned, which
    // moves no break, and each line that CSS justifies (not the last, not one ending in a soft break) gets the
    // `word-spacing` that takes its last character to the right edge, on spans around its text. A line without a space
    // stays at its start, as `inter-word` says. The spacing must not move a break: a paragraph in which a line no longer
    // fits after it (the line breaker sums widths its own way) goes back to left-aligned, every break where it was.
    // shortcut: horizontal left-to-right paragraphs only; an RTL or vertical one keeps Chromium's `auto`.
    {
        const INTER_WORD = /(?:^|; )text-justify: inter-word(?:;|$)/;
        // CSS word-separator characters, which word-spacing widens; trailing ones and zero-width ones end no line.
        const SEPARATOR = /^[  ]$/;
        const TRAILING = /^[\s​⁠　]$/;
        const paragraphs = Array.from(document.querySelectorAll<HTMLElement>(".text-paragraph")).filter(
            paragraph => INTER_WORD.test(paragraph.getAttribute("style") ?? "") && paragraph.style.direction !== "rtl" && paragraph.closest<HTMLElement>(".slide-text")?.dataset.textDirection === "horizontal",
        );

        if (paragraphs.length) {
            // Transforms move no layout: read lines in the boxes' own axes, where the text runs left to right even in a
            // turned or flipped box (a flipV box turns its text 180 degrees, #627).
            const flat = document.createElement("style");

            flat.textContent = "* { transform: none !important; }";
            document.head.append(flat);
            // Client rects are in zoomed px, word-spacing in CSS px; the line breaker rounds each piece up to a LayoutUnit.
            const zoom = parseFloat(getComputedStyle(document.documentElement).zoom) || 1;
            const LAYOUT_UNIT = 1 / 64;

            for (const paragraph of paragraphs) {
                paragraph.style.textAlign = "left";
            }

            const range = document.createRange();

            type Character = { node: Text; offset: number; left: number; right: number; text: string };
            const spreads: { paragraph: HTMLElement; characters: Character[]; spacing: number }[] = [];

            // All reads first, so the wrapping below does not lay the document out again for each line.
            for (const paragraph of paragraphs) {
                const style = getComputedStyle(paragraph);
                const edge = paragraph.getBoundingClientRect().right - (parseFloat(style.paddingRight) + parseFloat(style.borderRightWidth)) * zoom;
                const lines: Character[][] = [];
                const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);

                for (let node = walker.nextNode(); node; node = walker.nextNode()) {
                    if (!(node instanceof Text) || node.parentElement?.closest(".text-bullet")) {
                        continue;
                    }

                    for (let offset = 0; offset < node.length; offset++) {
                        range.setStart(node, offset);
                        range.setEnd(node, offset + 1);
                        const rect = range.getClientRects()[0];

                        if (!rect) {
                            continue;
                        }

                        const character = { node, offset, left: rect.left, right: rect.right, text: node.data.charAt(offset) };
                        const line = lines[lines.length - 1];
                        const previous = line?.[line.length - 1];

                        // Left to right, a line starts where the text moves back left.
                        if (!line || !previous || rect.left < previous.left - zoom) {
                            lines.push([character]);
                        } else {
                            line.push(character);
                        }
                    }
                }

                for (const [index, line] of lines.slice(0, -1).entries()) {
                    const last = line[line.length - 1];
                    const next = lines[index + 1]?.[0];

                    if (!last || !next) {
                        continue;
                    }

                    range.setStart(last.node, last.offset + 1);
                    range.setEnd(next.node, next.offset);
                    if (range.cloneContents().querySelector("br")) {
                        continue;
                    }

                    let end = line.length;

                    while (end > 0 && TRAILING.test(line[end - 1]?.text ?? "")) {
                        end--;
                    }

                    const content = line.slice(0, end);
                    const spaces = content.filter(({ text }) => SEPARATOR.test(text)).length;
                    const right = content[content.length - 1]?.right;

                    if (!spaces || right === undefined) {
                        continue;
                    }

                    const slack = edge - right - (spaces + 1) * LAYOUT_UNIT;

                    if (slack > 0) {
                        spreads.push({ paragraph, characters: content, spacing: slack / zoom / spaces });
                    }
                }
            }

            // Last line first: splitting a text node leaves the offsets of the lines before it valid.
            const spans = new Map<HTMLElement, HTMLElement[][]>();

            for (const { paragraph, characters, spacing } of spreads.reverse()) {
                const line: HTMLElement[] = [];

                spans.set(paragraph, [...(spans.get(paragraph) ?? []), line]);
                // One span per text node of the line, around the line's part of it.
                const parts = new Map<Text, { from: number; to: number }>();

                for (const { node, offset } of characters) {
                    const part = parts.get(node);

                    parts.set(node, { from: part?.from ?? offset, to: offset + 1 });
                }

                for (const [node, { from, to }] of parts) {
                    if (to < node.length) {
                        node.splitText(to);
                    }

                    const part = from > 0 ? node.splitText(from) : node;
                    const span = document.createElement("span");

                    span.style.wordSpacing = `${spacing}px`;
                    part.replaceWith(span);
                    span.append(part);
                    line.push(span);
                }
            }

            // A line still fits while each of its spans is one box on the line its first span is on: a wrapped part lands a
            // line height lower, where faces and raised runs move a box's middle by less than half its height.
            const middle = (rect: DOMRect) => rect.top + rect.height / 2;

            for (const lines of spans.values()) {
                const moved = lines.some(line => {
                    const first = line[0]?.getClientRects()[0];

                    return (
                        !first ||
                        line.some(span => {
                            const rects = span.getClientRects();
                            const rect = rects[0];

                            return rects.length !== 1 || !rect || Math.abs(middle(rect) - middle(first)) > Math.max(rect.height, first.height) / 2;
                        })
                    );
                });

                if (moved) {
                    for (const line of lines) {
                        for (const span of line) {
                            span.style.wordSpacing = "";
                        }
                    }
                }
            }

            flat.remove();
        }
    }

    // Strokes table borders as the profile draws them (docs/HTML_PROFILE.md, Table): one line per run of grid edges
    // with the same border, at the border's width, centred on the grid line. Chromium paints a CSS border as a filled
    // box floored to whole device px (a 0.5pt border prints 0.47pt high), which viewers anti-alias to a pale line,
    // where they draw a stroke crisp. The CSS borders keep their width, so the layout stays as it is, and turn
    // transparent; an SVG over the table strokes the lines on the edges of the cells, which a collapsed border centres
    // on the grid line. Read after the text has settled, since a row grows to its text.
    {
        const BORDER = /(?:^|; )border-(top|right|bottom|left): ([^ ;]+)px ([^ ;]+) ([^;]+)(?=;|$)/g;
        const SVG = "http://www.w3.org/2000/svg";
        const zoom = parseFloat(getComputedStyle(document.documentElement).zoom) || 1;

        type Side = { width: number; color: string; dash: string | undefined };
        for (const table of Array.from(document.querySelectorAll<HTMLTableElement>("table.slide-table"))) {
            const container = table.parentElement;

            if (!container) {
                continue;
            }

            // A transform moves no layout: measure without the object's rotation or flip, in CSS px from its corner.
            const transform = container.style.transform;

            container.style.transform = "none";
            const origin = container.getBoundingClientRect();
            const cells: { sides: Record<string, Side>; left: number; right: number; top: number; bottom: number }[] = [];

            for (const cell of Array.from(table.querySelectorAll<HTMLTableCellElement>("td"))) {
                const rect = cell.getBoundingClientRect();
                const sides: Record<string, Side> = {};

                for (const [, name = "", width, style, color = ""] of (cell.getAttribute("style") ?? "").matchAll(BORDER)) {
                    if (style === "none" || !(Number(width) > 0)) {
                        continue;
                    }

                    sides[name] = { width: Number(width), color, dash: cell.getAttribute(`data-border-${name}-dash-array`) ?? undefined };
                }

                const at = (value: number, from: number) => (value - from) / zoom;

                cells.push({ sides, left: at(rect.left, origin.left), right: at(rect.right, origin.left), top: at(rect.top, origin.top), bottom: at(rect.bottom, origin.top) });
            }

            container.style.transform = transform;
            // After every read: a style write between two reads makes Chromium recompute styles for the second.
            for (const cell of Array.from(table.querySelectorAll<HTMLTableCellElement>("td"))) {
                cell.style.borderColor = "transparent";
            }

            // The grid lines are the cells' edges; a key rounds away layout-unit noise. The first line is drawn on the
            // object's corner, where DrawingML starts the grid: the table offset centres only the widest border there.
            const key = (value: number) => Math.round(value * 64);
            const grid = (values: number[]) => {
                const sorted = Array.from(new Map(values.map(value => [key(value), value])).values()).sort((a, b) => a - b);
                const first = sorted[0] ?? 0;

                return { lines: sorted.map(value => value - first), index: new Map(sorted.map((value, index) => [key(value), index])) };
            };
            const xs = grid(cells.flatMap(cell => [cell.left, cell.right]));
            const ys = grid(cells.flatMap(cell => [cell.top, cell.bottom]));
            // Edge `line * steps + step` of each direction. The wider side wins an edge two cells share, as in CSS's
            // collapsing border model, and on a tie the first cell in document order keeps it.
            const horizontal: (Side | undefined)[] = [];
            const vertical: (Side | undefined)[] = [];
            const columns = xs.lines.length - 1;
            const rows = ys.lines.length - 1;
            const claim = (edges: (Side | undefined)[], slot: number, side: Side | undefined) => {
                const held = edges[slot];

                if (side && (!held || side.width > held.width)) {
                    edges[slot] = side;
                }
            };

            for (const { sides, left, right, top, bottom } of cells) {
                const [x0 = 0, x1 = 0, y0 = 0, y1 = 0] = [xs.index.get(key(left)), xs.index.get(key(right)), ys.index.get(key(top)), ys.index.get(key(bottom))];

                for (let column = x0; column < x1; column++) {
                    claim(horizontal, y0 * columns + column, sides.top);
                    claim(horizontal, y1 * columns + column, sides.bottom);
                }

                for (let row = y0; row < y1; row++) {
                    claim(vertical, x0 * rows + row, sides.left);
                    claim(vertical, x1 * rows + row, sides.right);
                }
            }

            const svg = document.createElementNS(SVG, "svg");

            svg.setAttribute("style", "position: absolute; left: 0px; top: 0px; width: 1px; height: 1px; overflow: visible;");
            /** The dashes of `side` from `from` to `to`: Skia prints a dashed stroke as filled outlines, so each dash is its own subpath. */
            const dashes = (side: Side, from: number, to: number): [number, number][] => {
                const pattern = (side.dash ?? "").split(/[\s,]+/).map(Number);
                let period = 0;

                for (const part of pattern) {
                    period += part;
                }

                if (!(period > 0) || pattern.some(part => !(part >= 0))) {
                    return [[from, to]];
                }

                const on: [number, number][] = [];

                for (let at = from, index = 0; at < to; index++) {
                    const length = pattern[index % pattern.length] ?? 0;

                    if (index % 2 === 0 && length > 0) {
                        on.push([at, Math.min(at + length, to)]);
                    }

                    at += length;
                }

                return on;
            };
            const strokeLines = (edges: (Side | undefined)[], lines: number[], steps: number[], segment: (at: number, from: number, to: number) => string) => {
                for (const [line, at] of lines.entries()) {
                    let run: { start: number; side: Side } | undefined;

                    for (let step = 0; step <= steps.length - 1; step++) {
                        const side = step < steps.length - 1 ? edges[line * (steps.length - 1) + step] : undefined;

                        if (run && side && side.width === run.side.width && side.color === run.side.color && side.dash === run.side.dash) {
                            continue;
                        }

                        if (run) {
                            const path = document.createElementNS(SVG, "path");

                            path.setAttribute(
                                "d",
                                dashes(run.side, steps[run.start] ?? 0, steps[step] ?? 0)
                                    .map(([from, to]) => segment(at, from, to))
                                    .join(" "),
                            );
                            path.setAttribute("fill", "none");
                            path.setAttribute("stroke", run.side.color);
                            path.setAttribute("stroke-width", String(run.side.width));
                            // Solid lines close the corners they share with the lines across them.
                            path.setAttribute("stroke-linecap", run.side.dash ? "butt" : "square");
                            svg.append(path);
                        }

                        run = side && { start: step, side };
                    }
                }
            };

            strokeLines(horizontal, ys.lines, xs.lines, (y, from, to) => `M ${from} ${y} H ${to}`);
            strokeLines(vertical, xs.lines, ys.lines, (x, from, to) => `M ${x} ${from} V ${to}`);
            if (svg.childElementCount) {
                table.after(svg);
            }
        }
    }

    // Draws the underline and strike of each run the print step marked (data-print-decoration) from the profile's
    // geometry (docs/HTML_PROFILE.md, Text, "Underline and strike geometry"), which CSS cannot state: Chromium draws its
    // own double, wavy and thick underlines, and one band per run (#396). A segment is the run's part on one line,
    // measured from its text's client rects; its baseline sits where a probe in the run's font puts it in the text's
    // content box. The pieces go into one SVG per text box, after its .text-content, in the box's CSS px, so they turn
    // with the box. Read after the text has settled: shrinking, raises and justification move it.
    // shortcut: horizontal text only; a vertical box keeps Chromium's decorations at the same em bands, until vertical
    // text is measured against LibreOffice.
    {
        const drawn: HTMLElement[] = [];

        for (const run of Array.from(document.querySelectorAll<HTMLElement>(".text-run[data-print-decoration]"))) {
            if (run.closest<HTMLElement>("[data-text-direction]")?.dataset.textDirection === "horizontal") {
                drawn.push(run);
            } else {
                run.style.textDecorationLine = run.dataset.printDecoration ?? "";
            }
        }

        if (drawn.length) {
            const SVG = "http://www.w3.org/2000/svg";
            const NUMBER = String.raw`-?(?:\d+\.?\d*|\.\d+)(?:e[-+]?\d+)?`;
            const EM = new RegExp(`^(${NUMBER})em$`);
            const STRIKE_BAND = new RegExp(`^(${NUMBER})em (${NUMBER})em$`);
            const PIECE = new RegExp(`^(line|wave|zigzag)\\(([^()]*)\\)$`);
            const PX = new RegExp(`^${NUMBER}px$`);
            // The smallest dash, height, half-wave, stroke or zigzag line the profile accepts (docs/HTML_PROFILE.md).
            const MIN_LENGTH = 0.01;
            const PLAIN = new RegExp(`^${NUMBER}$`);
            // The thinnest band the profile draws, in CSS px.
            const MIN_BAND = 0.5;
            // Widest gap, client px, between two underline segments of a line that one continues: rects of adjacent
            // runs meet up to layout rounding.
            const JOIN_PX = 0.05;
            const zoom = parseFloat(getComputedStyle(document.documentElement).zoom) || 1;

            type Piece =
                { kind: "line"; top: number; height: number; dashes: number[] } | { kind: "wave"; centre: number; amplitude: number; half: number; stroke: number } | { kind: "zigzag"; top: number; height: number; line: number; lineInBand: boolean };
            // A band in CSS px: its top below the baseline and its thickness.
            type Band = { offset: number; thickness: number };
            // A segment in client px, with its band; `joined` once an earlier one of its line continues over it.
            type Segment = { left: number; right: number; top: number; height: number; band: Band; joined: boolean };
            type Decoration = { pieces: Piece[]; color: string; segments: Segment[] };
            type Placed = { run: HTMLElement; svg: SVGSVGElement; origin: DOMRect; baseline: number; decorations: Decoration[] };
            const fail: (run: HTMLElement, what: string) => never = (run, what) => {
                throw new Error(`Cannot print '${run.closest<HTMLElement>("[data-object-name]")?.dataset.objectName ?? "unnamed object"}': a .text-run has ${what} (PdfLapas HTML profile, Text, Underline and strike geometry)`);
            };
            // data-underline or data-strike: pieces in thicknesses, each wave's last value a px stroke, each zigzag's a px
            // line or, unitless, one in thicknesses.
            const pieces = (run: HTMLElement, name: string): Piece[] => {
                const value = run.getAttribute(name) ?? "line(0 1)";
                const parsed: Piece[] = [];

                for (const written of value.match(/\S+\([^()]*\)/g) ?? []) {
                    const [, kind, list = ""] = written.match(PIECE) ?? [];
                    const parts = list.split(" ");
                    const last = parts.length - 1;
                    const values = parts.map(part => parseFloat(part));
                    const [a = NaN, b = NaN, c = NaN, d = NaN] = values;
                    const plain = parts.every((part, index) => (kind === "wave" && index === last ? PX : PLAIN).test(part) || (kind === "zigzag" && index === last && PX.test(part)));

                    if (kind === "line" && plain && parts.length >= 2 && parts.length % 2 === 0 && b >= MIN_LENGTH && values.slice(2).every(length => length >= MIN_LENGTH)) {
                        parsed.push({ kind, top: a, height: b, dashes: values.slice(2) });
                    } else if (kind === "wave" && plain && parts.length === 4 && b >= 0 && c >= MIN_LENGTH && d >= MIN_LENGTH) {
                        parsed.push({ kind, centre: a, amplitude: b, half: c, stroke: d });
                    } else if (kind === "zigzag" && plain && parts.length === 3 && b >= MIN_LENGTH && c >= MIN_LENGTH) {
                        parsed.push({ kind, top: a, height: b, line: c, lineInBand: PLAIN.test(parts[last] ?? "") });
                    } else {
                        fail(run, `${name} piece '${written}', which is not line(<top> <height> [<on> <off>]...), wave(<centre> <amplitude> <half> <w>px) or zigzag(<top> <height> <w>px|<w>)`);
                    }
                }

                // Every character belongs to a piece or a single space between two.
                if (!parsed.length || value.replace(/\S+\([^()]*\)/g, "").replace(/ /g, "") !== "" || value.includes("  ") || value.trim() !== value) {
                    fail(run, `${name}="${value}", which is not a space-separated list of line(), wave() and zigzag() pieces`);
                }

                return parsed;
            };
            const em = (run: HTMLElement, value: string, what: string): number => {
                const match = value.match(EM);

                return match ? Number(match[1]) : fail(run, `${what} '${value}', which is not a length in em`);
            };

            // Transforms move no layout: measure the runs and each box's SVG in the boxes' own axes.
            const flat = document.createElement("style");

            flat.textContent = "* { transform: none !important; }";
            document.head.append(flat);
            // One SVG per text box, in first, so its origin is read with the runs.
            const hosts = new Map<Element, SVGSVGElement>();

            for (const run of drawn) {
                const content = run.closest<HTMLElement>(".text-content");
                const host = content?.parentElement;

                if (content && host && !hosts.has(host)) {
                    const svg = document.createElementNS(SVG, "svg");

                    svg.setAttribute("style", "position: absolute; left: 0px; top: 0px; width: 1px; height: 1px; overflow: visible; pointer-events: none;");
                    // A mirrored .text-content (#627) mirrors its lines with it, about its own middle.
                    if (content.style.transform) {
                        svg.style.transformOrigin = `${content.offsetLeft + content.offsetWidth / 2}px 0px`;
                        svg.style.transform = content.style.transform;
                    }

                    host.append(svg);
                    hosts.set(host, svg);
                }
            }

            // A probe per font: a text and a zero-size inline-block on its baseline, which gives the baseline's place in the
            // text's content box (its client rect) as a fraction of that box's height.
            const probes = new Map<string, { probe: HTMLElement; text: Text; mark: HTMLElement; baseline: number }>();
            const fontKey = (run: HTMLElement) => `${run.style.fontFamily}|${run.style.fontSize}|${run.style.fontWeight}|${run.style.fontStyle}`;

            for (const run of drawn) {
                const key = fontKey(run);

                if (!probes.has(key)) {
                    const probe = document.createElement("div");
                    const text = document.createTextNode("x");
                    const mark = document.createElement("span");

                    probe.setAttribute("style", "position: absolute; left: 0px; top: 0px; white-space: pre; line-height: normal;");
                    probe.style.fontFamily = run.style.fontFamily;
                    probe.style.fontSize = run.style.fontSize;
                    probe.style.fontWeight = run.style.fontWeight;
                    probe.style.fontStyle = run.style.fontStyle;
                    mark.setAttribute("style", "display: inline-block; width: 0px; height: 0px; vertical-align: baseline;");
                    probe.append(text, mark);
                    document.body.append(probe);
                    probes.set(key, { probe, text, mark, baseline: 0 });
                }
            }

            const range = document.createRange();

            for (const probe of probes.values()) {
                range.selectNodeContents(probe.text);
                const box = range.getBoundingClientRect();

                probe.baseline = box.height > 0 ? (probe.mark.getBoundingClientRect().top - box.top) / box.height : 0;
            }

            /**
             * The segments of `run`'s text in client px, each with `band`: its text's rects merged where they share a line
             * (they overlap by more than half the smaller one's height).
             */
            const segments = (run: HTMLElement, band: Band): Segment[] => {
                const found: Segment[] = [];
                const walker = document.createTreeWalker(run, NodeFilter.SHOW_TEXT);

                for (let node = walker.nextNode(); node; node = walker.nextNode()) {
                    range.selectNodeContents(node);

                    for (const rect of Array.from(range.getClientRects())) {
                        if (!(rect.width > 0)) {
                            continue;
                        }

                        const last = found.at(-1);

                        if (last && Math.min(last.top + last.height, rect.bottom) - Math.max(last.top, rect.top) > Math.min(last.height, rect.height) / 2) {
                            last.left = Math.min(last.left, rect.left);
                            last.right = Math.max(last.right, rect.right);
                        } else {
                            found.push({ left: rect.left, right: rect.right, top: rect.top, height: rect.height, band, joined: false });
                        }
                    }
                }

                return found;
            };

            const placed: Placed[] = [];
            // The underline segments of data-underline-line-band runs, per paragraph, each with its baseline in client px
            // and its shape and colour.
            const lineBands = new Map<Element, { segment: Segment; baseline: number; look: string }[]>();

            // All reads first, so the drawing below does not lay the document out again for each run.
            for (const run of drawn) {
                const host = run.closest(".text-content")?.parentElement;
                const svg = host && hosts.get(host);
                const probe = probes.get(fontKey(run));
                const size = parseFloat(run.style.fontSize);
                const lines = run.dataset.printDecoration ?? "";

                if (!svg || !probe || !(size > 0)) {
                    fail(run, "no font-size in px or no .text-content around it");
                }

                const style = getComputedStyle(run);
                const band = (offset: number, thickness: number): Band => ({ offset: offset * size, thickness: Math.max(thickness * size, MIN_BAND) });
                const decorations: Decoration[] = [];

                if (lines.includes("underline")) {
                    const offset = em(run, run.style.textUnderlineOffset, "text-underline-offset");
                    const thickness = em(run, run.style.textDecorationThickness, "text-decoration-thickness");
                    const found = segments(run, band(offset, thickness));
                    const paragraph = run.closest(".text-paragraph");

                    if (paragraph && run.hasAttribute("data-underline-line-band")) {
                        const look = `${run.getAttribute("data-underline") ?? ""}|${style.textDecorationColor}`;
                        const entries = lineBands.get(paragraph) ?? [];

                        for (const segment of found) {
                            entries.push({ segment, baseline: segment.top + segment.height * probe.baseline, look });
                        }

                        lineBands.set(paragraph, entries);
                    }

                    decorations.push({ pieces: pieces(run, "data-underline"), color: style.textDecorationColor, segments: found });
                }

                if (lines.includes("line-through")) {
                    const written = run.dataset.strikeBand ?? "";
                    const [, offset = "", thickness = ""] = written.match(STRIKE_BAND) ?? fail(run, `data-strike-band="${written}", which is not '<offset>em <thickness>em'`);

                    decorations.push({ pieces: pieces(run, "data-strike"), color: style.color, segments: segments(run, band(Number(offset), Number(thickness))) });
                }

                placed.push({ run, svg, origin: svg.getBoundingClientRect(), baseline: probe.baseline, decorations });
            }

            flat.remove();
            for (const { probe } of probes.values()) {
                probe.remove();
            }

            // data-underline-line-band: the underlines on one line of a paragraph share one band, their bands averaged by
            // their segments' widths. A segment is on a line when it overlaps the line's first one by more than half.
            // Then a segment continues the one before it on its line when both have one shape and colour, sit on one
            // baseline and touch, so a dash pattern or zigzag runs on across runs (pdf-core's `join_line_band`).
            for (const found of lineBands.values()) {
                const lines: Segment[][] = [];

                for (const { segment } of found) {
                    const line = lines.find(([first]) => first && Math.min(first.top + first.height, segment.top + segment.height) - Math.max(first.top, segment.top) > Math.min(first.height, segment.height) / 2);

                    if (line) {
                        line.push(segment);
                    } else {
                        lines.push([segment]);
                    }
                }

                for (const line of lines) {
                    let width = 0;
                    let offset = 0;
                    let thickness = 0;

                    for (const { left, right, band } of line) {
                        width += right - left;
                        offset += band.offset * (right - left);
                        thickness += band.thickness * (right - left);
                    }

                    for (const segment of line) {
                        segment.band = { offset: offset / width, thickness: thickness / width };
                    }
                }

                let last: (typeof found)[number] | undefined;

                for (const entry of found) {
                    const { segment } = entry;

                    if (last && last.look === entry.look && Math.abs(last.baseline - entry.baseline) < JOIN_PX && Math.abs(segment.left - last.segment.right) < JOIN_PX) {
                        last.segment.right = Math.max(last.segment.right, segment.right);
                        segment.joined = true;
                    } else {
                        last = entry;
                    }
                }
            }

            const number = (value: number) => String(Number(value.toFixed(4)));

            for (const { run, svg, origin, baseline, decorations } of placed) {
                for (const { pieces: drawnPieces, color, segments: found } of decorations) {
                    let fill = "";

                    for (const segment of found) {
                        if (segment.joined) {
                            continue;
                        }

                        // The segment in the SVG's CSS px, P its band's top and t its thickness.
                        const start = (segment.left - origin.left) / zoom;
                        const end = (segment.right - origin.left) / zoom;
                        const p = (segment.top + segment.height * baseline - origin.top) / zoom + segment.band.offset;
                        const t = segment.band.thickness;

                        for (const piece of drawnPieces) {
                            if (piece.kind === "line") {
                                const top = p + piece.top * t;
                                const bottom = top + piece.height * t;
                                const period = piece.dashes.map(length => length * t);
                                // Without a dash pattern, or one with no length, the stripe is whole.
                                const dashed = period.some(length => length > 0);

                                for (let x = start, index = 0; x < end; index++) {
                                    const length = dashed ? (period[index % period.length] ?? 0) : end - start;

                                    if (index % 2 === 0 && length > 0) {
                                        fill += `M${number(x)} ${number(top)}H${number(Math.min(x + length, end))}V${number(bottom)}H${number(x)}Z`;
                                    }

                                    x += length;
                                }
                            } else if (piece.kind === "zigzag") {
                                // pdf-core's zigzag: 45-degree slopes, whole half-waves of `reach` from the top down first.
                                const top = p + piece.top * t;
                                const line = piece.lineInBand ? piece.line * t : piece.line;
                                const reach = Math.max(piece.height * t - line, line);
                                const halves = Math.max(Math.ceil((end - start) / reach), 1);
                                const lower: string[] = [];
                                const upper: string[] = [];

                                for (let index = 0; index <= halves; index++) {
                                    const edge = index % 2 === 0 ? top : top + reach;

                                    lower.push(`${number(start + reach * index)} ${number(edge + line)}`);
                                    upper.push(`${number(start + reach * index)} ${number(edge)}`);
                                }

                                fill += `M${number(start)} ${number(top)}L${lower.join("L")}L${upper.reverse().join("L")}Z`;
                            } else {
                                // Cubic half-waves about the centre line, as many as cover the segment: the first control
                                // point on the half-wave's start, the second at its middle, the amplitude off the line, up first.
                                const centre = p + piece.centre * t;
                                const half = piece.half * t;
                                const path = document.createElementNS(SVG, "path");
                                let d = `M${number(start)} ${number(centre)}`;

                                for (let x = start, index = 0; x < end; x += half, index++) {
                                    const peak = index % 2 === 0 ? centre - piece.amplitude * t : centre + piece.amplitude * t;

                                    d += `C${number(x)} ${number(centre)} ${number(x + half / 2)} ${number(peak)} ${number(x + half)} ${number(centre)}`;
                                }

                                path.setAttribute("d", d);
                                path.setAttribute("fill", "none");
                                path.setAttribute("stroke", color);
                                path.setAttribute("stroke-width", number(piece.stroke));
                                svg.append(path);
                            }
                        }
                    }

                    if (fill) {
                        const path = document.createElementNS(SVG, "path");

                        path.setAttribute("d", fill);
                        path.setAttribute("fill", color);
                        svg.append(path);
                    }
                }

                run.removeAttribute("data-print-decoration");
            }

            for (const svg of hosts.values()) {
                if (!svg.childElementCount) {
                    svg.remove();
                }
            }
        }
    }

    // The names of the images that did not decode.
    const failed: string[] = [];

    for (const image of Array.from(document.images)) {
        if (image.naturalWidth === 0 || image.naturalHeight === 0) {
            failed.push(image.alt || image.closest<HTMLElement>("[data-object-name]")?.dataset.objectName || "unnamed image");
        }
    }

    return failed;

    /** Resolves once `image` has loaded or failed. */
    function settled(image: HTMLImageElement): Promise<void> {
        if (image.complete) {
            return Promise.resolve();
        }

        return new Promise<void>(resolve => {
            image.addEventListener("load", () => resolve(), { once: true });
            image.addEventListener("error", () => resolve(), { once: true });
        });
    }
}

/** The Title (from <title>) and the other document properties from the meta tags, as the WASM reader reads them. */
function documentInfo(): DocumentInfo {
    const meta = (name: string) => document.querySelector<HTMLMetaElement>(`meta[name="${name}"]`)?.content;

    return {
        Title: document.title,
        Author: meta("author"),
        Subject: meta("description"),
        Company: meta("pdflapas-company"),
        Revision: meta("pdflapas-revision"),
        Keywords: meta("keywords"),
        Creator: meta("pdflapas-creator"),
        Producer: meta("pdflapas-producer"),
        CreationDate: meta("pdflapas-creation-date"),
        ModDate: meta("pdflapas-modification-date"),
    };
}

/** Points each external link at `placeholder` and its index (see `withLinkUris`), and returns the links' URLs and objects. */
function placeLinkPlaceholders(placeholder: string): { href: string; object: string }[] {
    const links: { href: string; object: string }[] = [];

    for (const [index, link] of Array.from(document.querySelectorAll('a[href]:not([href^="#"])')).entries()) {
        const href = link.getAttribute("href") ?? "";

        link.setAttribute("href", `${placeholder}${index}`);
        const slide = link.closest(".slide-container")?.id.replace(/^slide-/, "") ?? "?";
        const name = link.closest<HTMLElement>("[data-object-name]")?.dataset.objectName;

        links.push({ href, object: name === undefined ? `an object on slide ${slide}` : `'${name}' on slide ${slide}` });
    }

    return links;
}

/**
 * The run faces Chromium has no font for, with the objects asking for them and the face drawn instead. A family is
 * available when putting it before a generic family changes a sample's width against that generic alone (two generics
 * of different widths, so a font as wide as one of them still counts). Fontconfig's metric aliases (Arial -> Liberation
 * Sans, Calibri -> Carlito) count as the face: Chromium draws them for it. The width checks run once per distinct
 * font-family list; per run it reads the inline style. `document.fonts.check()` cannot tell: it is true for any
 * family that is not a web font.
 */
function missingFonts(): MissingFont[] {
    const sample = "mmmmmmmmmmlli WQ@#0";
    const context = document.createElement("canvas").getContext("2d");

    if (!context) {
        throw new Error("Cannot check the deck's fonts: this Chromium build has no 2D canvas");
    }

    const width = (fontFamily: string) => {
        context.font = `16px ${fontFamily}`;

        return context.measureText(sample).width;
    };
    const generics = ["monospace", "serif"].map(generic => ({ generic, alone: width(generic) }));
    const available = (family: string) => generics.some(({ generic, alone }) => width(`${family}, ${generic}`) !== alone);
    // Chromium serializes a family as an identifier when it is one, else as a string with `"` and `\` escaped.
    const unquote = (family: string) => (family.startsWith('"') ? family.slice(1, -1).replace(/\\(.)/g, "$1") : family);
    // A font-family list -> its missing first face and the family drawn instead, or null when the face is available.
    const lists = new Map<string, { face: string; usedFace: string | undefined } | null>();
    const missing = new Map<string, { usedFace: string | undefined; objects: Set<string> }>();

    for (const run of Array.from(document.querySelectorAll<HTMLElement>(".text-run"))) {
        const list = run.style.fontFamily;

        if (!list) {
            continue;
        } // Out-of-profile HTML; the profile gives every run a face.

        if (!lists.has(list)) {
            const [face = "", ...fallbacks] = (list.match(/"(?:[^"\\]|\\.)*"|[^,\s][^,]*/g) ?? []).map(family => family.trim());
            const used = available(face) ? null : (fallbacks.find(family => available(family)) ?? "");

            lists.set(list, used === null ? null : { face: unquote(face), usedFace: used ? unquote(used) : undefined });
        }

        const found = lists.get(list);

        if (!found) {
            continue;
        }

        const entry = missing.get(found.face) ?? { usedFace: found.usedFace, objects: new Set<string>() };

        missing.set(found.face, entry);
        const object = run.closest<HTMLElement>(".slide-element");
        const slide = run.closest(".slide-container")?.id.replace(/^slide-/, "") ?? "?";
        const index = object?.parentElement ? Array.prototype.indexOf.call(object.parentElement.children, object) : -1;

        entry.objects.add(object?.dataset.objectName === undefined ? `slide ${slide} #${index}` : `slide ${slide} '${object.dataset.objectName}'`);
    }

    return Array.from(missing, ([face, { usedFace, objects }]) => ({ face, objects: [...objects].sort(), ...(usedFace ? { usedFace } : {}) }));
}
