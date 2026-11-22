// Migration check (#408): compares two PDFs of the same deck, for example the old PPTX-to-PDF pipeline's and
// PdfLapas's, and reports what a reviewer looks for: page count and size, text changes in reading order, lines
// that break after a different word, words that moved, shapes and images with no match, and a pixel score.
//
//   bun run compare:pdfs reference.pdf actual.pdf [--json] [--move-tolerance=1] [--shape-tolerance=2] [--pixel-threshold=0.12] [--diff-dir=<dir>]
//
// With --diff-dir, each page that differs gets an ImageMagick diff image there, page-<n>-diff.png.
//
// Exits 1 when a difference exceeds the tolerances, 2 on a usage error. Reads the PDFs with poppler
// (pdftotext, pdftocairo, pdftoppm) and ImageMagick, the tools tests/visual already needs.
import { copyFile, mkdir, mkdtemp, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { parseArgs } from "node:util";
import { PAGE_THRESHOLD } from "../tests/visual/thresholds";

/** A word moved when its origin moved more than this, in pt (#408 asks for 1pt). */
export const MOVE_TOLERANCE_PT = 1;

/** A shape or image matches when each edge of its box is within this many pt of the other PDF's. */
export const SHAPE_TOLERANCE_PT = 2;

/** Pixel score limit per page: the visual parity page limit (normalized RMSE at 96 dpi). */
export const PIXEL_THRESHOLD = PAGE_THRESHOLD;

/**
 * Two pdftotext words on one line with a gap below this fraction of their height are one word. pdftotext splits a
 * letter-spaced word (charSpacing) into fragments where PowerPoint draws each glyph apart, and keeps it whole where
 * Chromium draws one run; a space in Arial is about 0.24 of the word height, so real word breaks stay.
 * shortcut: spacing wider than this still splits a word; cluster the gaps of a line if a deck needs that.
 */
const FRAGMENT_GAP_RATIO = 0.15;

/** The text report lists this many items of one kind per page, then a count; the JSON lists all of them. */
const LISTED_ITEMS = 8;

/** Resolution of the rasterized pages for the pixel score, as tests/visual/compare.ts renders them. */
const RASTER_DPI = 96;

const REQUIRED_TOOLS = ["pdftotext", "pdftocairo", "pdftoppm", "magick"] as const;

export type CompareOptions = {
    moveTolerance: number;
    shapeTolerance: number;
    pixelThreshold: number;
    /** Folder that receives `page-<n>-diff.png` for each page that differs. */
    diffDir?: string;
};

export type Box = { kind: "shape" | "image"; x0: number; y0: number; x1: number; y1: number };

export type PageReport = {
    page: number;
    size: { reference: [number, number]; actual: [number, number] } | { same: [number, number] };
    removedWords: string[];
    addedWords: string[];
    /** A word both PDFs draw that ends a line in one of them only. */
    lineBreaks: { word: string; endsLineIn: "reference" | "actual" }[];
    movedWords: { word: string; dx: number; dy: number }[];
    /** A box of the reference that the actual PDF draws at another place with the same size. */
    movedBoxes: { box: Box; dx: number; dy: number }[];
    missingBoxes: Box[];
    extraBoxes: Box[];
    pixelScore: number;
    /** The diff image written for this page under `diffDir`. */
    diffImage?: string;
};

export type CompareReport = {
    pages: { reference: number; actual: number };
    options: CompareOptions;
    pageReports: PageReport[];
    passed: boolean;
};

/** `x` is the first glyph's origin; `y` is the baseline once `alignBaselines` found it, the box's bottom before. */
export type Word = { text: string; x: number; y: number; height: number; endsLine: boolean };

/** A glyph origin, on the baseline, in pt from the top left. */
export type Point = { x: number; y: number };

export type Page = { width: number; height: number; words: Word[] };

type Matrix = [number, number, number, number, number, number];

const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

/** Below this a derivative coefficient counts as zero, in pt: far below the shape tolerance. */
const CUBIC_EPSILON = 1e-9;

const PAGE_PATTERN = /<page width="([\d.]+)" height="([\d.]+)">([\s\S]*?)<\/page>/g;
const LINE_PATTERN = /<line [^>]*>([\s\S]*?)<\/line>/g;
const WORD_PATTERN = /<word xMin="([\d.-]+)" yMin="([\d.-]+)" xMax="([\d.-]+)" yMax="([\d.-]+)">([^<]*)<\/word>/g;
const SUBPATH_PATTERN = /(?=M)/;
const ENTITY_PATTERN = /&(amp|lt|gt|quot|apos);/g;
const ENTITIES: Readonly<Record<string, string>> = { amp: "&", lt: "<", gt: ">", quot: '"', apos: "'" };
const TAG_PATTERN = /<(\/?)(g|path|use|image)\b([^>]*?)(\/?)>/g;
// cairo writes path data with absolute M, L, C and Z only.
const PATH_TOKEN_PATTERN = /[MLCZ]|-?\d*\.?\d+(?:e-?\d+)?/g;
/** Numbers each path command takes per segment. */
const PATH_ARGUMENTS: Readonly<Record<string, number>> = { M: 2, L: 2, C: 6, Z: 0 };
const PNG_PAGE_PATTERN = /-(\d+)\.png$/;
const ID_ATTRIBUTE = / id="([^"]+)"/;
const WIDTH_ATTRIBUTE = / width="([\d.]+)"/;
const HEIGHT_ATTRIBUTE = / height="([\d.]+)"/;
const PATH_DATA_ATTRIBUTE = / d="([^"]*)"/;
const X_ATTRIBUTE = / x="([\d.e-]+)"/;
const Y_ATTRIBUTE = / y="([\d.e-]+)"/;
const HREF_ATTRIBUTE = /xlink:href="#([^"]+)"/;
const MATRIX_ATTRIBUTE = / transform="matrix\(([^)]*)\)"/;
const RMSE_PATTERN = /\((\d*\.?\d+(?:e-?\d+)?)\)/;

async function run(command: string[], acceptedExitCodes = [0]): Promise<string> {
    const child = Bun.spawn(command, { stdout: "pipe", stderr: "pipe" });
    const [exitCode, stdout, stderr] = await Promise.all([child.exited, new Response(child.stdout).text(), new Response(child.stderr).text()]);

    if (!acceptedExitCodes.includes(exitCode)) {
        throw new Error(`${command.join(" ")} failed (${exitCode})\n${stderr || stdout}`);
    }

    return `${stdout}\n${stderr}`.trim();
}

/** The pages of a PDF with their words in reading order, from `pdftotext -bbox-layout` output. */
export function readPages(xhtml: string): Page[] {
    const pages: Page[] = [];

    for (const [, width, height, body] of xhtml.matchAll(PAGE_PATTERN)) {
        const words: Word[] = [];

        for (const [, line] of (body ?? "").matchAll(LINE_PATTERN)) {
            const start = words.length;
            let previousEnd = -Infinity;

            for (const [, x0, y0, x1, y1, raw] of (line ?? "").matchAll(WORD_PATTERN)) {
                const text = (raw ?? "").replace(ENTITY_PATTERN, (_, name: string) => ENTITIES[name] ?? name);
                const previous = words[words.length - 1];

                if (previous && words.length > start && Number(x0) - previousEnd < FRAGMENT_GAP_RATIO * (Number(y1) - Number(y0))) {
                    previous.text += text;
                } else {
                    words.push({ text, x: Number(x0), y: Number(y1), height: Number(y1) - Number(y0), endsLine: false });
                }

                previousEnd = Number(x1);
            }

            const last = words[words.length - 1];

            if (last && words.length > start) {
                last.endsLine = true;
            }
        }

        pages.push({ width: Number(width), height: Number(height), words });
    }

    return pages;
}

/**
 * Moves each word's `y` from pdftotext's box bottom (the baseline plus the font's descent, which differs between
 * fonts) to its baseline: the origin of the glyph closest to the word's start that sits inside the word's box.
 * shortcut: a word with no glyph origin there (text cairo draws as paths) keeps its box bottom; match on the text
 * matrices of the content stream if a deck hits that.
 */
export function alignBaselines(words: Word[], origins: Point[]): void {
    for (const word of words) {
        let best: Point | undefined;

        for (const origin of origins) {
            const inside = origin.y <= word.y && origin.y >= word.y - word.height && Math.abs(origin.x - word.x) <= word.height / 2;

            if (inside && (!best || Math.abs(origin.x - word.x) < Math.abs(best.x - word.x))) {
                best = origin;
            }
        }

        if (best) {
            word.y = best.y;
        }
    }
}

/**
 * The boxes of the filled or stroked paths and the images on one page, and the origins of its glyphs, in pt from
 * the top left.
 */
export function readGraphics(svg: string, width: number, height: number): { boxes: Box[]; origins: Point[] } {
    const defsEnd = svg.indexOf("</defs>");
    const imageSizes = new Map<string, [number, number]>();

    for (const [, , , attributes] of svg.slice(0, Math.max(defsEnd, 0)).matchAll(TAG_PATTERN)) {
        const id = attributes?.match(ID_ATTRIBUTE)?.[1];
        const w = attributes?.match(WIDTH_ATTRIBUTE)?.[1];
        const h = attributes?.match(HEIGHT_ATTRIBUTE)?.[1];

        if (id && w && h) {
            imageSizes.set(id, [Number(w), Number(h)]);
        }
    }

    const boxes: Box[] = [];
    const origins: Point[] = [];
    const stack: Matrix[] = [IDENTITY];

    for (const [, closing, tag, attributes = "", selfClosing] of svg.slice(Math.max(defsEnd, 0)).matchAll(TAG_PATTERN)) {
        const parent = stack[stack.length - 1] ?? IDENTITY;

        if (tag === "g") {
            if (closing) {
                stack.pop();
            } else if (!selfClosing) {
                stack.push(multiply(parent, attributes));
            }

            continue;
        }

        const matrix = multiply(parent, attributes);
        const [a, b, c, d, e, f] = matrix;
        const at = (x: number, y: number): Point => ({ x: a * x + c * y + e, y: b * x + d * y + f });
        // Each subpath's extreme points in page space: its line ends, and for a Bézier its end and the points where it
        // turns, so the box is the curve's, not its control points'.
        const subpaths: Point[][] = [];
        let kind: Box["kind"] = "shape";

        if (tag === "path") {
            // Each subpath is its own box: one pipeline draws a chart's gridlines as one path, another as one path each.
            for (const subpath of (attributes.match(PATH_DATA_ATTRIBUTE)?.[1] ?? "").split(SUBPATH_PATTERN)) {
                const points: Point[] = [];
                const numbers: number[] = [];
                let command = "M";

                for (const [token] of subpath.matchAll(PATH_TOKEN_PATTERN)) {
                    if (token in PATH_ARGUMENTS) {
                        command = token;
                        numbers.length = 0;
                        continue;
                    }

                    numbers.push(Number(token));
                    if (numbers.length < (PATH_ARGUMENTS[command] ?? 0)) {
                        continue;
                    }

                    const [x1 = 0, y1 = 0, x2 = 0, y2 = 0, x3 = 0, y3 = 0] = numbers;
                    const start = points[points.length - 1];

                    numbers.length = 0;
                    if (command === "C" && start) {
                        const end = at(x3, y3);

                        points.push(...cubicTurns(start, at(x1, y1), at(x2, y2), end), end);
                    } else {
                        points.push(at(x1, y1));
                    }
                }

                subpaths.push(points);
            }
        } else {
            // A <use> of an image source is an image; any other <use> is a glyph, whose origin gives a word's baseline.
            const size = imageSizes.get(attributes.match(HREF_ATTRIBUTE)?.[1] ?? "");

            if (tag === "image") {
                continue;
            }

            if (!size) {
                origins.push(at(Number(attributes.match(X_ATTRIBUTE)?.[1] ?? 0), Number(attributes.match(Y_ATTRIBUTE)?.[1] ?? 0)));
                continue;
            }

            kind = "image";
            subpaths.push([at(0, 0), at(size[0], size[1])]);
        }

        for (const points of subpaths) {
            // A lone moveto (cairo ends a closed path with one) draws nothing.
            if (points.length < 2) {
                continue;
            }

            let box: Box = { kind, x0: Infinity, y0: Infinity, x1: -Infinity, y1: -Infinity };

            for (const { x, y } of points) {
                box = { kind, x0: Math.min(box.x0, x), y0: Math.min(box.y0, y), x1: Math.max(box.x1, x), y1: Math.max(box.y1, y) };
            }

            // The page background, which one pipeline paints and another leaves blank, is the pixel score's business.
            const coversPage = box.x0 <= 1 && box.y0 <= 1 && box.x1 >= width - 1 && box.y1 >= height - 1;

            if (!coversPage) {
                boxes.push(box);
            }
        }
    }

    return { boxes, origins };
}

/**
 * The points strictly inside a cubic Bézier where it turns in x or in y: where a coordinate's derivative,
 * 3(at² + bt + c), is zero for t in (0, 1). The box of the curve is the box of these and its two ends.
 */
export function cubicTurns(p0: Point, p1: Point, p2: Point, p3: Point): Point[] {
    const turns: Point[] = [];

    for (const axis of ["x", "y"] as const) {
        const a = -p0[axis] + 3 * p1[axis] - 3 * p2[axis] + p3[axis];
        const b = 2 * (p0[axis] - 2 * p1[axis] + p2[axis]);
        const c = p1[axis] - p0[axis];
        const discriminant = b * b - 4 * a * c;
        // A linear derivative (a = 0) has one root; a quadratic has two, or none below a zero discriminant.
        const roots = Math.abs(a) < CUBIC_EPSILON ? (Math.abs(b) < CUBIC_EPSILON ? [] : [-c / b]) : discriminant < 0 ? [] : [(-b + Math.sqrt(discriminant)) / (2 * a), (-b - Math.sqrt(discriminant)) / (2 * a)];

        for (const t of roots) {
            if (t > 0 && t < 1) {
                const u = 1 - t;

                turns.push({
                    x: u * u * u * p0.x + 3 * u * u * t * p1.x + 3 * u * t * t * p2.x + t * t * t * p3.x,
                    y: u * u * u * p0.y + 3 * u * u * t * p1.y + 3 * u * t * t * p2.y + t * t * t * p3.y,
                });
            }
        }
    }

    return turns;
}

/** The parent matrix times the element's own `transform="matrix(...)"`, if it has one. */
function multiply(parent: Matrix, attributes: string): Matrix {
    const own = attributes.match(MATRIX_ATTRIBUTE)?.[1];

    if (!own) {
        return parent;
    }

    const [a = 1, b = 0, c = 0, d = 1, e = 0, f = 0] = own.split(",").map(Number);
    const [pa, pb, pc, pd, pe, pf] = parent;

    return [pa * a + pc * b, pb * a + pd * b, pa * c + pc * d, pb * c + pd * d, pa * e + pc * f + pe, pb * e + pd * f + pf];
}

/** The boxes of `from` with no box of the same kind within `tolerance` on every edge in `to`. */
function unmatched(from: Box[], to: Box[], tolerance: number): Box[] {
    const result: Box[] = [];

    for (const box of from) {
        let found = false;

        for (const other of to) {
            if (other.kind === box.kind && Math.abs(other.x0 - box.x0) <= tolerance && Math.abs(other.y0 - box.y0) <= tolerance && Math.abs(other.x1 - box.x1) <= tolerance && Math.abs(other.y1 - box.y1) <= tolerance) {
                found = true;
                break;
            }
        }

        if (!found) {
            result.push(box);
        }
    }

    return result;
}

/** Compares `actual` against `reference`, page by page. */
export async function comparePdfs(reference: string, actual: string, options: CompareOptions): Promise<CompareReport> {
    for (const tool of REQUIRED_TOOLS) {
        if (!Bun.which(tool)) {
            throw new Error(`Comparing PDFs requires '${tool}' on PATH (poppler-utils and ImageMagick)`);
        }
    }

    const [referenceText, actualText] = await Promise.all([run(["pdftotext", "-bbox-layout", "-enc", "UTF-8", reference, "-"]), run(["pdftotext", "-bbox-layout", "-enc", "UTF-8", actual, "-"])]);
    const referencePages = readPages(referenceText);
    const actualPages = readPages(actualText);
    const shared = Math.min(referencePages.length, actualPages.length);
    // Normalized RMSE of each page pair, rasterized at RASTER_DPI, and its diff image when a diff folder is named.
    const directory = await mkdtemp(join(tmpdir(), "pdflapas-compare-"));
    const scores: number[] = [];
    const pageReports: PageReport[] = [];
    let passed = referencePages.length === actualPages.length;

    try {
        await Promise.all([run(["pdftoppm", "-png", "-r", String(RASTER_DPI), "-l", String(shared), reference, join(directory, "reference")]), run(["pdftoppm", "-png", "-r", String(RASTER_DPI), "-l", String(shared), actual, join(directory, "actual")])]);
        const files = (await readdir(directory)).sort((a, b) => Number(a.match(PNG_PAGE_PATTERN)?.[1]) - Number(b.match(PNG_PAGE_PATTERN)?.[1]));

        for (const file of files) {
            if (!file.startsWith("reference-")) {
                continue;
            }

            const diff = options.diffDir ? join(directory, `diff-${scores.length + 1}.png`) : "null:";
            // Without the date chunks, the same pages give a byte-identical diff image.
            const output = await run(["magick", "compare", "-metric", "RMSE", "-define", "png:exclude-chunks=date,time", join(directory, file), join(directory, file.replace("reference-", "actual-")), diff], [0, 1]);
            const score = output.match(RMSE_PATTERN)?.[1];

            if (!score) {
                throw new Error(`Could not parse the ImageMagick metric: ${output}`);
            }

            scores.push(Number(score));
        }

        for (let index = 0; index < shared; index++) {
            const ref = referencePages[index];
            const act = actualPages[index];

            if (!ref || !act) {
                continue;
            }

            const [referenceSvg, actualSvg] = await Promise.all([run(["pdftocairo", "-svg", "-f", String(index + 1), "-l", String(index + 1), reference, "-"]), run(["pdftocairo", "-svg", "-f", String(index + 1), "-l", String(index + 1), actual, "-"])]);
            const referenceGraphics = readGraphics(referenceSvg, ref.width, ref.height);
            const actualGraphics = readGraphics(actualSvg, act.width, act.height);

            alignBaselines(ref.words, referenceGraphics.origins);
            alignBaselines(act.words, actualGraphics.origins);

            // Longest common subsequence of the words: what is left out is removed or added text, and the pairs it
            // keeps are the same word in both PDFs, whose line ends and positions are compared.
            const n = ref.words.length;
            const m = act.words.length;
            const lengths = new Uint32Array((n + 1) * (m + 1));

            for (let i = n - 1; i >= 0; i--) {
                for (let j = m - 1; j >= 0; j--) {
                    lengths[i * (m + 1) + j] = ref.words[i]?.text === act.words[j]?.text ? (lengths[(i + 1) * (m + 1) + j + 1] ?? 0) + 1 : Math.max(lengths[(i + 1) * (m + 1) + j] ?? 0, lengths[i * (m + 1) + j + 1] ?? 0);
                }
            }

            const report: PageReport = {
                page: index + 1,
                size: Math.abs(ref.width - act.width) <= options.shapeTolerance && Math.abs(ref.height - act.height) <= options.shapeTolerance ? { same: [ref.width, ref.height] } : { reference: [ref.width, ref.height], actual: [act.width, act.height] },
                removedWords: [],
                addedWords: [],
                lineBreaks: [],
                movedWords: [],
                movedBoxes: [],
                missingBoxes: [],
                extraBoxes: [],
                pixelScore: scores[index] ?? 1,
            };
            const pairs: [Word, Word][] = [];
            const removed: Word[] = [];
            const added: Word[] = [];
            let i = 0;
            let j = 0;

            while (i < n || j < m) {
                const refWord = ref.words[i];
                const actWord = act.words[j];

                if (refWord && actWord && refWord.text === actWord.text) {
                    pairs.push([refWord, actWord]);
                    i++;
                    j++;
                } else if (refWord && (!actWord || (lengths[(i + 1) * (m + 1) + j] ?? 0) >= (lengths[i * (m + 1) + j + 1] ?? 0))) {
                    removed.push(refWord);
                    i++;
                } else if (actWord) {
                    added.push(actWord);
                    j++;
                }
            }

            // The LCS follows reading order, which two PDFs can disagree on (a chart's axis labels drawn before or after
            // its title). A word left out on both sides is the same word drawn elsewhere: pair it with the nearest twin.
            for (const refWord of removed) {
                let nearest = -1;

                for (const [index, actWord] of added.entries()) {
                    const best = added[nearest];

                    if (actWord.text === refWord.text && (!best || Math.hypot(actWord.x - refWord.x, actWord.y - refWord.y) < Math.hypot(best.x - refWord.x, best.y - refWord.y))) {
                        nearest = index;
                    }
                }

                const twin = added[nearest];

                if (twin) {
                    pairs.push([refWord, twin]);
                    added.splice(nearest, 1);
                } else {
                    report.removedWords.push(refWord.text);
                }
            }

            for (const actWord of added) {
                report.addedWords.push(actWord.text);
            }

            for (const [refWord, actWord] of pairs) {
                if (refWord.endsLine !== actWord.endsLine) {
                    report.lineBreaks.push({ word: refWord.text, endsLineIn: refWord.endsLine ? "reference" : "actual" });
                }

                const dx = actWord.x - refWord.x;
                const dy = actWord.y - refWord.y;

                if (Math.hypot(dx, dy) > options.moveTolerance) {
                    report.movedWords.push({ word: refWord.text, dx: Number(dx.toFixed(2)), dy: Number(dy.toFixed(2)) });
                }
            }

            const referenceBoxes = referenceGraphics.boxes;
            const actualBoxes = actualGraphics.boxes;
            const extraBoxes = unmatched(actualBoxes, referenceBoxes, options.shapeTolerance);

            // A box with no match whose twin (same kind and size) is among the extra boxes moved rather than vanished.
            for (const box of unmatched(referenceBoxes, actualBoxes, options.shapeTolerance)) {
                const twin = extraBoxes.findIndex(other => other.kind === box.kind && Math.abs(other.x1 - other.x0 - (box.x1 - box.x0)) <= options.shapeTolerance && Math.abs(other.y1 - other.y0 - (box.y1 - box.y0)) <= options.shapeTolerance);
                const moved = extraBoxes[twin];

                if (moved) {
                    report.movedBoxes.push({ box, dx: Number((moved.x0 - box.x0).toFixed(2)), dy: Number((moved.y0 - box.y0).toFixed(2)) });
                    extraBoxes.splice(twin, 1);
                } else {
                    report.missingBoxes.push(box);
                }
            }

            const differs =
                !("same" in report.size) ||
                report.removedWords.length + report.addedWords.length + report.lineBreaks.length + report.movedWords.length + report.movedBoxes.length + report.missingBoxes.length + extraBoxes.length > 0 ||
                report.pixelScore > options.pixelThreshold;

            if (differs) {
                passed = false;
            }

            if (differs && options.diffDir) {
                const diffImage = join(options.diffDir, `page-${index + 1}-diff.png`);

                await mkdir(options.diffDir, { recursive: true });
                await copyFile(join(directory, `diff-${index + 1}.png`), diffImage);
                pageReports.push({ ...report, extraBoxes, diffImage });
            } else {
                pageReports.push({ ...report, extraBoxes });
            }
        }
    } finally {
        await rm(directory, { recursive: true, force: true });
    }

    return { pages: { reference: referencePages.length, actual: actualPages.length }, options, pageReports, passed };
}

/** The report as a short text summary, one block per page that differs. */
export function formatReport(report: CompareReport): string {
    const lines: string[] = [];
    const sample = (items: string[]) => `${items.slice(0, LISTED_ITEMS).join(", ")}${items.length > LISTED_ITEMS ? `, ... (${items.length})` : ""}`;
    const box = (b: Box) => `${b.kind} [${b.x0.toFixed(1)}, ${b.y0.toFixed(1)}, ${b.x1.toFixed(1)}, ${b.y1.toFixed(1)}]`;

    lines.push(`Pages: reference ${report.pages.reference}, actual ${report.pages.actual}${report.pages.reference === report.pages.actual ? "" : "  DIFFERENT"}`);

    for (const page of report.pageReports) {
        const details: string[] = [];

        if (!("same" in page.size)) {
            details.push(`  size: reference ${page.size.reference.join(" x ")} pt, actual ${page.size.actual.join(" x ")} pt`);
        }

        if (page.removedWords.length > 0) {
            details.push(`  text only in reference: ${sample(page.removedWords)}`);
        }

        if (page.addedWords.length > 0) {
            details.push(`  text only in actual: ${sample(page.addedWords)}`);
        }

        if (page.movedWords.length > 0) {
            details.push(`  ${page.movedWords.length} words moved more than ${report.options.moveTolerance}pt: ${sample(page.movedWords.map(word => `"${word.word}" (${word.dx}, ${word.dy})`))}`);
        }

        // One line per item, the first LISTED_ITEMS of each kind, then how many more the JSON lists.
        const listed: [string, string[]][] = [
            ["line breaks", page.lineBreaks.map(lineBreak => `line break after "${lineBreak.word}" only in ${lineBreak.endsLineIn}`)],
            ["moved boxes", page.movedBoxes.map(moved => `${box(moved.box)} moved by (${moved.dx}, ${moved.dy})`)],
            ["missing boxes", page.missingBoxes.map(missing => `${box(missing)} in reference has no match in actual`)],
            ["extra boxes", page.extraBoxes.map(extra => `${box(extra)} in actual has no match in reference`)],
        ];

        for (const [name, items] of listed) {
            for (const item of items.slice(0, LISTED_ITEMS)) {
                details.push(`  ${item}`);
            }

            if (items.length > LISTED_ITEMS) {
                details.push(`  ... ${items.length - LISTED_ITEMS} more ${name} (${items.length} in all)`);
            }
        }

        const pixel = `pixel RMSE ${page.pixelScore.toFixed(4)}${page.pixelScore > report.options.pixelThreshold ? ` > ${report.options.pixelThreshold}` : ""}`;

        if (page.diffImage) {
            details.push(`  diff image: ${page.diffImage}`);
        }

        lines.push(`Page ${page.page}: ${details.length === 0 && page.pixelScore <= report.options.pixelThreshold ? "same" : "differs"}, ${pixel}`);
        lines.push(...details);
    }

    lines.push(report.passed ? "PASS" : "FAIL");

    return lines.join("\n");
}

if (import.meta.main) {
    const { values, positionals } = parseArgs({
        allowPositionals: true,
        options: {
            json: { type: "boolean", default: false },
            "move-tolerance": { type: "string", default: String(MOVE_TOLERANCE_PT) },
            "shape-tolerance": { type: "string", default: String(SHAPE_TOLERANCE_PT) },
            "pixel-threshold": { type: "string", default: String(PIXEL_THRESHOLD) },
            "diff-dir": { type: "string" },
        },
    });
    const [reference, actual] = positionals;
    const tolerances = { moveTolerance: Number(values["move-tolerance"]), shapeTolerance: Number(values["shape-tolerance"]), pixelThreshold: Number(values["pixel-threshold"]) };

    if (!reference || !actual || positionals.length !== 2 || values["diff-dir"] === "" || !Object.values(tolerances).every(value => Number.isFinite(value) && value >= 0)) {
        console.error("Usage: bun run compare:pdfs <reference.pdf> <actual.pdf> [--json] [--move-tolerance=<pt>] [--shape-tolerance=<pt>] [--pixel-threshold=<rmse>] [--diff-dir=<dir>]");
        process.exit(2);
    }

    const report = await comparePdfs(reference, actual, { ...tolerances, ...(values["diff-dir"] ? { diffDir: values["diff-dir"] } : {}) });

    console.log(values.json ? JSON.stringify(report, null, 2) : formatReport(report));
    process.exit(report.passed ? 0 : 1);
}
