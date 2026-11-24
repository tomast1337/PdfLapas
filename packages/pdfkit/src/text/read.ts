// Reads a text box of the profile (docs/HTML_PROFILE.md "Text") into the values the layout needs. Every inline
// property and data-* attribute the renderer does not draw is reported through `report`, never dropped silently.
import { type Color, color, cssStrings, em, px, readStyle, type Style, sides, take, takePx } from "#src/css";
import { attribute, children, type Element, hasAttribute, hasClass, isElement, textContent } from "#src/dom";
import { type Effect, glowEffect } from "#src/effect/index";
import type { FaceRequest } from "#src/fonts";

/**
 * A piece of a decoration around its band (docs/HTML_PROFILE.md "Underline and strike geometry"), lengths in band
 * thicknesses T down from the band's top: a stripe (cut into `dashes`, on and off in turn), a stroked wave, a zigzag.
 */
export type DecorationPiece =
    | { readonly kind: "line"; readonly top: number; readonly height: number; readonly dashes: readonly number[] }
    | { readonly kind: "wave"; readonly centre: number; readonly amplitude: number; readonly half: number; readonly stroke: number }
    /** `line` px tall, or `line` thicknesses when `lineInBand`. */
    | { readonly kind: "zigzag"; readonly top: number; readonly height: number; readonly line: number; readonly lineInBand: boolean };

export type Decoration = {
    readonly color: Color;
    readonly pieces: readonly DecorationPiece[];
    /** The band the HTML states, in em: its top `top` below the baseline, `thickness` tall. Without it, the face's (draw.ts). */
    readonly band?: { readonly top: number; readonly thickness: number };
    /**
     * `data-underline-line-band`: the underline takes its line's averaged band, and continues the one before it on its
     * line when both have one `look` (their `data-underline` and colour as written).
     */
    readonly lineBand?: { readonly look: string };
};

export type Run = {
    /** The run's text; U+200B (an empty run) and U+2060 (no break) take part in breaking only. */
    readonly text: string;
    readonly face: FaceRequest;
    readonly size: number;
    readonly color: Color;
    readonly letterSpacing: number;
    readonly kerning: boolean;
    readonly underline?: Decoration;
    readonly strike?: Decoration;
    /** `top` of a raised or lowered run (super/subscript, baseline offset), px down. */
    readonly shift: number;
    readonly highlight?: Color;
    readonly outline?: { readonly width: number; readonly color: Color };
    readonly link?: { readonly href?: string; readonly slide?: number };
    readonly softBreakBefore: boolean;
    readonly synthesis: Synthesis;
    /** The run's glow: its `text-shadow` entry that `data-text-shadow` names `glow`. */
    readonly glow?: Effect;
};

/** How a run is drawn when its face has no bold or italic style (`data-synthetic-bold`, `data-synthetic-italic`). */
export type Synthesis = {
    /** The stroke drawn around the filled glyphs, as a fraction of the font size. */
    readonly boldStroke: number;
    /** The shear to the right per unit of height above the baseline. */
    readonly italicShear: number;
};

export type Bullet = {
    readonly text: string;
    readonly face: FaceRequest;
    readonly size: number;
    readonly color: Color;
};

export type Align = "left" | "center" | "right" | "justify";

export type TabStop = {
    readonly position: number;
    readonly alignment: "left" | "right" | "center" | "decimal";
};

export type Paragraph = {
    readonly runs: readonly Run[];
    readonly align: Align;
    /** `direction: rtl`: each line's words run right to left. */
    readonly rtl: boolean;
    readonly lineHeight: number;
    readonly baseline: number;
    readonly markBaseline: number;
    readonly breakBaseline: number;
    readonly spaceBefore: number;
    readonly spaceAfter: number;
    readonly tabSize: number;
    /** `data-tab-stops`, px from the text area's left edge, in order. */
    readonly tabStops: readonly TabStop[];
    /** `padding-left`: every line starts here. */
    readonly indent: number;
    /** `text-indent`: the first line, and its bullet, start this far from `indent`. */
    readonly textIndent: number;
    readonly bullet?: Bullet;
};

export type Fit = { readonly kind: "none" | "resize" } | { readonly kind: "shrink"; readonly step: number; readonly maxSteps: number; readonly tolerance: number };

export type TextBox = {
    readonly paragraphs: readonly Paragraph[];
    readonly padding: readonly [number, number, number, number];
    readonly anchor: "flex-start" | "center" | "flex-end";
    readonly wrap: boolean;
    readonly fit: Fit;
    /** `.text-content` `transform: scaleX(-1)`: the text mirrors back inside a box flipped one way. */
    readonly mirrored: boolean;
};

export type Report = (what: string) => void;

/** A text box read, or one whose direction the renderer does not draw (already reported). */
export type ReadTextBox = { readonly kind: "text"; readonly box: TextBox } | { readonly kind: "vertical" };

const ALIGNS: ReadonlySet<string> = new Set(["left", "center", "right", "justify"]);
const TAB_ALIGNS: ReadonlySet<string> = new Set(["left", "right", "center", "decimal"]);
const DIRECTIONS: ReadonlySet<string> = new Set(["ltr", "rtl"]);
/** Letters of the right-to-left scripts (Hebrew, Arabic, Syriac, Thaana, NKo, …, and their presentation forms). */
const RTL_LETTER = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/;
/** A letter of any other script: a strong left-to-right character. */
const LTR_LETTER = /(?![\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF])\p{L}/u;
const ANCHORS: ReadonlySet<string> = new Set(["flex-start", "center", "flex-end"]);
/** A decoration without `data-underline` or `data-strike`: one stripe filling its band. */
const BAND_PIECES: readonly DecorationPiece[] = Object.freeze([{ kind: "line", top: 0, height: 1, dashes: [] }]);
/** The pieces of a `data-underline` or `data-strike` value, a closing parenthesis then spaces apart. */
const PIECE_SEPARATOR = /(?<=\))\s+/;
const PIECE = /^(line|wave|zigzag)\(([^()]*)\)$/;
const PX_ARGUMENT = /^(.*)px$/;
/** The smallest decoration length the profile accepts (docs/HTML_PROFILE.md): below it a segment would take more pieces than are worth drawing. */
const MIN_DECORATION_LENGTH = 0.01;
/**
 * Chromium's synthetic styles, the default for HTML that does not state them: pdf-core's
 * CHROMIUM_SYNTHETIC_BOLD_STROKE_EM and CHROMIUM_SYNTHETIC_ITALIC_SHEAR, measured in Chromium's PDF output.
 */
const BROWSER_SYNTHESIS: Synthesis = Object.freeze({ boldStroke: 1 / 32, italicShear: 0.25 });
/**
 * What a U+2011 in a run with each `data-non-breaking-hyphen` value is drawn as: U+002D or U+2010, then U+2060 WORD
 * JOINER, which keeps it unbreakable.
 */
const NON_BREAKING_HYPHEN_DRAWN: ReadonlyMap<string, string> = new Map([
    ["hyphen-minus", "-\u2060"],
    ["hyphen", "\u2010\u2060"],
]);
const NON_BREAKING_HYPHEN = /\u2011/g;
/** The value of each underline geometry property that the renderer draws (or the em band, read before these). */
const UNDERLINE_FROM_FONT = Object.freeze([
    ["text-decoration-thickness", "from-font"],
    ["text-underline-position", "from-font"],
    ["text-decoration-skip-ink", "none"],
] as const);
/** Layout CSS that `toHtml()` writes for Chromium and that states no drawn value of its own. */
const LAYOUT_ONLY: Readonly<Record<string, ReadonlySet<string>>> = Object.freeze({
    content: new Set(["width"]),
    bullet: new Set(["display", "width"]),
});

/** Reports the properties left in `style` and the data-* attributes not in `read`. */
export function reportUnread(element: Element, style: Style, read: readonly string[], report: Report, ignored: ReadonlySet<string> = new Set()): void {
    for (const [property, value] of style) {
        if (!ignored.has(property)) {
            report(`CSS property '${property}: ${value}' on <${element.tagName}> is not drawn`);
        }
    }

    for (const { name } of element.attrs) {
        if (name.startsWith("data-") && !read.includes(name)) {
            report(`attribute '${name}' on <${element.tagName}> is not drawn`);
        }
    }
}

/** `value`, which a missing or unreadable input gives as NaN. */
/** A `<n>px` property the profile requires: missing or unreadable is an invalid document. */
function requiredPx(style: Style, property: string, what: string): number {
    const length = px(take(style, property, ""));

    if (length.kind === "invalid" || !Number.isFinite(length.value)) {
        throw new Error(`invalid PdfLapas HTML: ${what} is missing or not a number`);
    }

    return length.value;
}

/** A numeric attribute the profile requires: missing or unreadable is an invalid document. */
function requiredNumber(element: Element, name: string, what: string): number {
    if (!hasAttribute(element, name)) {
        throw new Error(`invalid PdfLapas HTML: ${what} is missing or not a number`);
    }

    const value = Number(attribute(element, name, ""));

    if (!Number.isFinite(value)) {
        throw new Error(`invalid PdfLapas HTML: ${what} is missing or not a number`);
    }

    return value;
}

function requiredColor(value: string, what: string): Color {
    const parsed = color(value);

    if (parsed.kind === "invalid") {
        throw new Error(`invalid PdfLapas HTML: ${what} '${value}' is missing or not a color the renderer reads (#rgb, #rrggbb, rgb(), rgba())`);
    }

    return parsed.value;
}

/** `value` as a color, or reported as one the renderer does not read. */
function readColor(value: string, what: string, report: Report): { kind: "color"; color: Color } | { kind: "invalid" } {
    const parsed = color(value);

    if (parsed.kind === "invalid") {
        report(`${what} '${value}' is not a color the renderer reads`);

        return { kind: "invalid" };
    }

    return { kind: "color", color: parsed.value };
}

/** A `data-underline` or `data-strike` value as its pieces, or what is wrong with it. */
function readPieces(value: string): { kind: "pieces"; pieces: DecorationPiece[] } | { kind: "invalid"; reason: string } {
    const pieces: DecorationPiece[] = [];

    for (const written of value.trim().split(PIECE_SEPARATOR)) {
        const match = PIECE.exec(written);
        const [kind, list] = [match?.[1], match?.[2]];

        if (kind === undefined || list === undefined) {
            return { kind: "invalid", reason: `'${written}' is not line(...), wave(...) or zigzag(...)` };
        }

        // The arguments as numbers; the last of a wave is a px length, of a zigzag a px length or a plain number.
        const words = list.trim().split(/\s+/);
        const numbers: number[] = [];
        const lineInBand = kind === "zigzag" && !PX_ARGUMENT.test(words.at(-1) ?? "");

        for (const [index, word] of words.entries()) {
            const pxLength = kind !== "line" && !lineInBand && index === words.length - 1;
            const digits = pxLength ? PX_ARGUMENT.exec(word)?.[1] : word;
            const number = digits === undefined || digits === "" ? Number.NaN : Number(digits);

            if (!Number.isFinite(number)) {
                return { kind: "invalid", reason: `'${written}' has '${word}' where a ${pxLength ? "px length" : "number"} goes` };
            }

            numbers.push(number);
        }

        const [first = 0, second = 0, third = 0, fourth = 0] = numbers;

        if (kind === "line") {
            const dashes = numbers.slice(2);

            if (numbers.length < 2 || dashes.length % 2 !== 0 || second < MIN_DECORATION_LENGTH || dashes.some(length => length < MIN_DECORATION_LENGTH)) {
                return { kind: "invalid", reason: `'${written}' is not line(<top> <height>) with a height and an even count of dash lengths of at least ${MIN_DECORATION_LENGTH}` };
            }

            pieces.push({ kind: "line", top: first, height: second, dashes });
        } else if (kind === "wave") {
            if (numbers.length !== 4 || second < 0 || third < MIN_DECORATION_LENGTH || fourth < MIN_DECORATION_LENGTH) {
                return { kind: "invalid", reason: `'${written}' is not wave(<centre> <amplitude> <half> <w>px) with an amplitude of at least 0, and a half and width of at least ${MIN_DECORATION_LENGTH}` };
            }

            pieces.push({ kind: "wave", centre: first, amplitude: second, half: third, stroke: fourth });
        } else {
            if (numbers.length !== 3 || second < MIN_DECORATION_LENGTH || third < MIN_DECORATION_LENGTH) {
                return { kind: "invalid", reason: `'${written}' is not zigzag(<top> <height> <w>px|<w>) with a height and width of at least ${MIN_DECORATION_LENGTH}` };
            }

            pieces.push({ kind: "zigzag", top: first, height: second, line: third, lineInBand });
        }
    }

    return { kind: "pieces", pieces };
}

/** The pieces of decoration attribute `name`: one stripe filling the band when it is omitted, or malformed (reported). */
function decorationPieces(element: Element, name: string, report: Report): readonly DecorationPiece[] {
    if (!hasAttribute(element, name)) {
        return BAND_PIECES;
    }

    const value = attribute(element, name, "");
    const read = readPieces(value);

    if (read.kind === "invalid") {
        report(`attribute '${name}' '${value}' is not drawn: ${read.reason} (drawn as line(0 1))`);

        return BAND_PIECES;
    }

    return read.pieces;
}

function readFace(element: Element, style: Style): FaceRequest {
    const value = take(style, "font-family", "");
    const families = cssStrings(value);
    const [family, ...fallbacks] = families.kind === "parsed" ? families.value : [];

    if (family === undefined) {
        throw new Error(`invalid PdfLapas HTML: font-family '${value}' on <${element.tagName}> is not one or more quoted families separated by ', '`);
    }

    const lang = attribute(element, "lang", "");

    return { family, fallbacks, ...(lang ? { lang } : {}), bold: take(style, "font-weight", "") === "bold", italic: take(style, "font-style", "") === "italic" };
}

/** `text-shadow` entries: commas inside `rgba(...)` do not separate them. */
const SHADOW_ENTRY = /[^,(]+(?:\([^)]*\)[^,(]*)*/g;

/**
 * The text box of `box` (`.slide-text` or `.table-cell-text`), whose style map the caller already read; the text
 * properties are taken out of it. Vertical, after a report, for any direction but horizontal.
 */
export function readTextBox(box: Element, style: Style, report: Report): ReadTextBox {
    const direction = attribute(box, "data-text-direction", "horizontal");
    const anchor = take(style, "align-items", "center");
    const whiteSpace = take(style, "white-space", "pre-wrap");

    style.delete("overflow-wrap");
    style.delete("direction"); // each paragraph repeats it, and pdf-core reads it there only
    const paddingSides = sides(take(style, "padding", ""));
    const padding: TextBox["padding"] = paddingSides.kind === "parsed" ? paddingSides.value : [0, 0, 0, 0];

    if (direction !== "horizontal") {
        report(`${direction} text is not drawn`);

        return { kind: "vertical" };
    }

    if (!ANCHORS.has(anchor)) {
        report(`align-items '${anchor}' is not drawn`);
    }

    const content = children(box).find(child => hasClass(child, "text-content"));

    if (!content) {
        throw new Error("invalid PdfLapas HTML: a text box has no .text-content");
    }

    const contentStyle = readStyle(content);
    const transform = take(contentStyle, "transform", "none");

    if (transform !== "none" && transform !== "scaleX(-1)") {
        throw new Error(`invalid PdfLapas HTML: a .text-content's transform is '${transform}', expected scaleX(-1)`);
    }

    reportUnread(content, contentStyle, [], report, LAYOUT_ONLY.content);
    const paragraphs: Paragraph[] = [];

    for (const paragraphElement of children(content)) {
        const paragraphStyle = readStyle(paragraphElement);
        const align = take(paragraphStyle, "text-align", "left");

        if (!ALIGNS.has(align)) {
            report(`text-align '${align}' is not drawn`);
        }

        const paragraphDirection = take(paragraphStyle, "direction", "ltr");

        if (!DIRECTIONS.has(paragraphDirection)) {
            throw new Error(`invalid PdfLapas HTML: a paragraph's direction is '${paragraphDirection}', expected ltr or rtl`);
        }

        const rtl = paragraphDirection === "rtl";
        const lineHeight = requiredPx(paragraphStyle, "line-height", "a paragraph's line-height");
        const baseline = requiredNumber(paragraphElement, "data-line-baseline", "a paragraph's data-line-baseline");
        const optionalBaseline = (name: string) => (hasAttribute(paragraphElement, name) ? requiredNumber(paragraphElement, name, `a paragraph's ${name}`) : baseline);
        const parts = children(paragraphElement);
        const bulletElement = parts.find(child => hasClass(child, "text-bullet"));
        const runs: Run[] = [];

        for (const child of parts) {
            if (!hasClass(child, "text-run")) {
                continue;
            }

            const runStyle = readStyle(child);
            const runColor = requiredColor(take(runStyle, "color", ""), "a run's color");
            const [hasHref, hasSlide] = [hasAttribute(child, "href"), hasAttribute(child, "data-slide")];
            const position = take(runStyle, "position", "relative");
            const shift = takePx(runStyle, "top", 0);

            if (position !== "relative") {
                report(`position '${position}' on a run is not drawn`);
            }

            const hasHighlight = runStyle.has("background-color");
            const highlightValue = take(runStyle, "background-color", "");
            const highlight = hasHighlight ? readColor(highlightValue, "background-color", report) : { kind: "none" as const };
            const letterSpacing = takePx(runStyle, "letter-spacing", 0);
            // The run's glow from `text-shadow`, whose entries `data-text-shadow` names in order. The `shadow` entry is the
            // object's shadow, which slide.ts draws from the box's `box-shadow` (pdf-core skips it the same way).
            const hasShadow = runStyle.has("text-shadow");
            const shadow = take(runStyle, "text-shadow", "");
            let glow: Effect | undefined;

            if (hasAttribute(child, "data-shadow-inset")) {
                report("inner shadow not rendered");
            }

            if (hasShadow) {
                const names = attribute(child, "data-text-shadow", "").split(/\s+/).filter(Boolean);
                const entries: string[] = [];

                for (const entry of shadow.match(SHADOW_ENTRY) ?? []) {
                    const trimmed = entry.trim();

                    if (trimmed) {
                        entries.push(trimmed);
                    }
                }

                if (names.length !== entries.length) {
                    report(`text-shadow '${shadow}' has ${entries.length} entries but data-text-shadow names ${names.length}; not drawn`);
                } else {
                    for (const [index, name] of names.entries()) {
                        const entry = entries[index] ?? "";

                        if (name === "shadow") {
                            continue;
                        }

                        if (name !== "glow") {
                            report(`text-shadow entry '${entry}' named '${name}' (not glow or shadow) is not drawn`);
                            continue;
                        }

                        const effect = glowEffect(entry);

                        if (effect.kind === "invalid") {
                            report(effect.report);
                        } else {
                            glow = effect.kind === "effect" ? effect.effect : undefined;
                        }
                    }
                }
            }

            const face = readFace(child, runStyle);
            const size = requiredPx(runStyle, "font-size", "a run's font-size");
            const kerning = take(runStyle, "font-kerning", "") !== "none";
            // The run's underline and strikethrough.
            const decorationLines = take(runStyle, "text-decoration-line", "").split(/\s+/);
            const cssStyle = take(runStyle, "text-decoration-style", "solid");
            const hasDecorationColor = runStyle.has("text-decoration-color");
            const decorationColorValue = take(runStyle, "text-decoration-color", "");
            const decorationColorRead = hasDecorationColor ? readColor(decorationColorValue, "text-decoration-color", report) : { kind: "none" as const };
            const decorationColor = decorationColorRead.kind === "color" ? decorationColorRead.color : runColor;
            // The underline's band in em: `text-underline-offset` and `text-decoration-thickness`, without
            // `text-underline-position`; both taken from the style. None without an offset; any other form is reported.
            const hasOffset = runStyle.has("text-underline-offset");
            const offset = take(runStyle, "text-underline-offset", "");
            let band: Decoration["band"];

            if (hasOffset) {
                // Read, not taken: the thickness stays for the from-font check below unless it belongs to the band.
                const [offsetEm, thicknessEm] = [em(offset), em(runStyle.get("text-decoration-thickness") ?? "")];

                if (offsetEm.kind === "invalid" || thicknessEm.kind === "invalid" || thicknessEm.value <= 0 || runStyle.has("text-underline-position")) {
                    report(`CSS property 'text-underline-offset: ${offset}' is not drawn: an underline band is '<n>em' with 'text-decoration-thickness' a positive '<n>em' and no 'text-underline-position'`);
                } else {
                    runStyle.delete("text-decoration-thickness");
                    band = { top: offsetEm.value, thickness: thicknessEm.value };
                }
            }

            // Without an em band the underline takes the face's post band (draw.ts); any other value is reported.
            for (const [property, drawn] of UNDERLINE_FROM_FONT) {
                const declared = runStyle.has(property);
                const value = take(runStyle, property, "");

                if (declared && value !== drawn) {
                    report(`CSS property '${property}: ${value}' is not drawn (drawn as ${drawn})`);
                }
            }

            const doubleStrike = hasAttribute(child, "data-double-strike");
            const underlined = decorationLines.includes("underline");
            // The underline's CSS style is Chromium's fallback: data-underline states what is drawn. A double strike
            // takes text-decoration-style, the underline's own style then being in data-underline-style.
            const underlineStyle = doubleStrike ? attribute(child, "data-underline-style", "solid") : cssStyle;

            if (underlined && underlineStyle !== "solid" && !hasAttribute(child, "data-underline")) {
                report(`underline style '${underlineStyle}' without 'data-underline' is not drawn (drawn as line(0 1))`);
            }

            // The strike's band in em: data-strike-band="<top below the baseline>em <thickness>em".
            const strikeBandValue = attribute(child, "data-strike-band", "");
            const [strikeTop = "", strikeThickness = "", ...strikeRest] = strikeBandValue.trim().split(/\s+/);
            const [strikeTopEm, strikeThicknessEm] = [em(strikeTop), em(strikeThickness)];
            let strikeBand: Decoration["band"];

            if (hasAttribute(child, "data-strike-band")) {
                if (strikeTopEm.kind === "invalid" || strikeThicknessEm.kind === "invalid" || strikeThicknessEm.value <= 0 || strikeRest.length > 0) {
                    report(`attribute 'data-strike-band' '${strikeBandValue}' is not drawn: it is '<top>em <thickness>em' with a positive thickness (drawn with the face's strikeout band)`);
                } else {
                    strikeBand = { top: strikeTopEm.value, thickness: strikeThicknessEm.value };
                }
            }

            // text-decoration-color is the underline's (docs/HTML_PROFILE.md): the strike is in the text colour, as
            // PowerPoint and pdf-core draw it (OOXML gives a strike no colour of its own).
            const underline: Decoration = {
                color: decorationColor,
                pieces: decorationPieces(child, "data-underline", report),
                ...(band ? { band } : {}),
                ...(hasAttribute(child, "data-underline-line-band") ? { lineBand: { look: `${attribute(child, "data-underline", "")}|${decorationColor.rgb.join(",")},${decorationColor.alpha}` } } : {}),
            };
            const strike: Decoration = { color: runColor, pieces: decorationPieces(child, "data-strike", report), ...(strikeBand ? { band: strikeBand } : {}) };
            const decorations: Pick<Run, "underline" | "strike"> = {
                ...(underlined ? { underline } : {}),
                ...(decorationLines.includes("line-through") || doubleStrike ? { strike } : {}),
            };
            // `-webkit-text-stroke: <width>px <color>`; any other value is reported.
            const hasOutline = runStyle.has("-webkit-text-stroke");
            const stroke = take(runStyle, "-webkit-text-stroke", "");
            let outline: Run["outline"];

            if (hasOutline) {
                // A value without a space (a width or a colour alone) is not the profile's form, as pdf-core refuses it.
                const space = stroke.indexOf(" ");
                const width = px(space === -1 ? "" : stroke.slice(0, space));
                const outlineColor = color(space === -1 ? "" : stroke.slice(space + 1));

                if (width.kind === "invalid" || outlineColor.kind === "invalid") {
                    report(`-webkit-text-stroke '${stroke}' is not drawn`);
                } else {
                    outline = { width: width.value, color: outlineColor.value };
                }
            }

            // How the run is drawn when its face has no bold or italic style (`data-synthetic-bold`, `data-synthetic-italic`).
            const syntheticNumber = (name: string, fallback: number) => {
                if (!hasAttribute(child, name)) {
                    return fallback;
                }

                const value = attribute(child, name, "");
                const parsed = Number(value);

                if (value.trim() === "" || !Number.isFinite(parsed)) {
                    throw new Error(`invalid PdfLapas HTML: attribute '${name}' '${value}' is not a number`);
                }

                return parsed;
            };
            const boldStroke = syntheticNumber("data-synthetic-bold", BROWSER_SYNTHESIS.boldStroke);

            if (boldStroke < 0) {
                throw new Error(`invalid PdfLapas HTML: attribute 'data-synthetic-bold' ${boldStroke} is negative`);
            }

            const run: Run = {
                text: textContent(child),
                face,
                size,
                color: runColor,
                letterSpacing,
                kerning,
                ...decorations,
                shift,
                ...(highlight.kind === "color" && highlight.color.alpha > 0 ? { highlight: highlight.color } : {}),
                ...(hasOutline ? { outline } : {}),
                ...(hasHref || hasSlide ? { link: { ...(hasHref ? { href: attribute(child, "href", "") } : {}), ...(hasSlide ? { slide: Number(attribute(child, "data-slide", "")) } : {}) } } : {}),
                softBreakBefore: hasAttribute(child, "data-soft-break-before"),
                synthesis: { boldStroke, italicShear: syntheticNumber("data-synthetic-italic", BROWSER_SYNTHESIS.italicShear) },
                ...(glow ? { glow } : {}),
            };
            const read = [
                "data-double-strike",
                "data-underline-style",
                "data-underline",
                "data-underline-line-band",
                "data-strike",
                "data-strike-band",
                "data-vertical-align",
                "data-empty",
                "data-soft-break-before",
                "data-slide",
                "data-text-shadow",
                "data-shadow-inset",
            ];

            reportUnread(child, runStyle, [...read, "data-synthetic-bold", "data-synthetic-italic", "data-non-breaking-hyphen"], report);
            // The run's text in pieces of one `letter-spacing` each: its bare text at the run's, each word span
            // (`<span style="letter-spacing: ...">`, text only) at its own. U+2011 becomes U+002D or U+2010, then U+2060,
            // under `data-non-breaking-hyphen`.
            const hasHyphen = hasAttribute(child, "data-non-breaking-hyphen");
            const hyphenValue = attribute(child, "data-non-breaking-hyphen", "");
            const hyphen = NON_BREAKING_HYPHEN_DRAWN.get(hyphenValue);

            if (hasHyphen && hyphen === undefined) {
                throw new Error(`invalid PdfLapas HTML: attribute 'data-non-breaking-hyphen' '${hyphenValue}' is not hyphen-minus or hyphen`);
            }

            const hyphenated = (value: string) => (hyphen === undefined ? value : value.replace(NON_BREAKING_HYPHEN, hyphen));
            const pieces: { text: string; letterSpacing: number; word: boolean }[] = [];

            for (const node of child.childNodes) {
                if (!isElement(node)) {
                    const value = "value" in node ? node.value : "";
                    const last = pieces.at(-1);

                    // Bare text pieces next to each other share the run's spacing.
                    if (last && !last.word) {
                        last.text += hyphenated(value);
                    } else {
                        pieces.push({ text: hyphenated(value), letterSpacing, word: false });
                    }

                    continue;
                }

                const spanStyle = readStyle(node);
                const spacing = px(take(spanStyle, "letter-spacing", ""));

                if (node.tagName !== "span" || node.attrs.length !== 1 || spacing.kind === "invalid" || spanStyle.size || node.childNodes.some(isElement)) {
                    throw new Error(`invalid PdfLapas HTML: <${node.tagName}> inside a run; a run holds text and word spans, <span style="letter-spacing: <n>px;"> with text only`);
                }

                pieces.push({ text: hyphenated(textContent(node)), letterSpacing: spacing.value, word: true });
            }

            // One run per piece of one letter-spacing; only the first starts after a soft break.
            const drawnPieces = pieces.filter(piece => piece.text);

            if (!drawnPieces.length) {
                runs.push(run);
            } else {
                for (const [index, piece] of drawnPieces.entries()) {
                    runs.push({ ...run, text: piece.text, letterSpacing: piece.letterSpacing, softBreakBefore: run.softBreakBefore && index === 0 });
                }
            }
        }

        for (const child of parts) {
            if (!hasClass(child, "text-run") && !hasClass(child, "text-bullet") && child.tagName !== "br") {
                report(`<${child.tagName}> in a paragraph is not drawn`);
            }
        }

        const text = runs.map(run => run.text).join("");

        // shortcut: no Unicode bidi algorithm, so a paragraph's words run in its direction; port one (or take pdf-core's
        // parley order) when a deck mixes directions.
        if (rtl ? LTR_LETTER.test(text) : RTL_LETTER.test(text)) {
            report(`${rtl ? "left-to-right words in a right-to-left" : "right-to-left words in a left-to-right"} paragraph are ordered in the paragraph's direction`);
        }

        if (rtl && text.includes("\t")) {
            report("tab stops in right-to-left paragraphs measured from the left");
        }

        const markBaseline = optionalBaseline("data-mark-baseline");
        const breakBaseline = optionalBaseline("data-break-baseline");
        // `data-tab-stops="<position>:<left|right|center|decimal>,…"`.
        const tabStops: TabStop[] = [];

        if (hasAttribute(paragraphElement, "data-tab-stops")) {
            for (const stop of attribute(paragraphElement, "data-tab-stops", "").split(",")) {
                const [stopPosition, alignment = ""] = stop.split(":");
                const parsed = Number(stopPosition);

                if (!stopPosition?.trim() || !Number.isFinite(parsed) || !TAB_ALIGNS.has(alignment)) {
                    throw new Error(`invalid PdfLapas HTML: attribute 'data-tab-stops': '${stop}' is not '<position>:<left|right|center|decimal>'`);
                }

                tabStops.push({ position: parsed, alignment: alignment as TabStop["alignment"] }); // `as`: checked against TAB_ALIGNS
            }
        }

        let bullet: Bullet | undefined;

        if (bulletElement) {
            const bulletStyle = readStyle(bulletElement);

            bullet = {
                text: textContent(bulletElement),
                face: readFace(bulletElement, bulletStyle),
                size: requiredPx(bulletStyle, "font-size", "a bullet's font-size"),
                color: requiredColor(take(bulletStyle, "color", ""), "a bullet's color"),
            };
            reportUnread(bulletElement, bulletStyle, [], report, LAYOUT_ONLY.bullet);
        }

        const paragraph: Paragraph = {
            runs,
            align: ALIGNS.has(align) ? (align as Align) : "left", // `as`: checked against ALIGNS
            rtl,
            lineHeight,
            baseline,
            markBaseline,
            breakBaseline,
            spaceBefore: takePx(paragraphStyle, "margin-top", 0),
            spaceAfter: takePx(paragraphStyle, "margin-bottom", 0),
            tabSize: takePx(paragraphStyle, "tab-size", 0),
            tabStops,
            indent: takePx(paragraphStyle, "padding-left", 0),
            textIndent: takePx(paragraphStyle, "text-indent", 0),
            ...(bullet ? { bullet } : {}),
        };
        const bulletAttributes = ["data-bullet-level", "data-bullet-type", "data-bullet-indent", "data-bullet-start", "data-bullet-number-type"];

        reportUnread(paragraphElement, paragraphStyle, ["data-line-baseline", "data-mark-baseline", "data-break-baseline", "data-tab-stops", ...bulletAttributes], report);
        paragraphs.push(paragraph);
    }

    const fitValue = attribute(box, "data-text-fit", "none");
    let fit: Fit;

    if (fitValue !== "shrink") {
        fit = { kind: fitValue === "resize" ? "resize" : "none" };
    } else {
        const number = (name: string) => requiredNumber(box, name, `a shrink box's ${name}`);

        fit = { kind: "shrink", step: number("data-text-fit-step"), maxSteps: number("data-text-fit-max-steps"), tolerance: number("data-text-fit-tolerance") };
    }

    return {
        kind: "text",
        box: {
            paragraphs,
            padding,
            anchor: ANCHORS.has(anchor) ? (anchor as TextBox["anchor"]) : "center", // `as`: checked against ANCHORS
            wrap: whiteSpace !== "pre",
            fit,
            mirrored: transform === "scaleX(-1)",
        },
    };
}
