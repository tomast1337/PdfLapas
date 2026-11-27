import type PptxGenJS from "pptxgenjs";
import { PPTX_DEFAULTS } from "#src/defaults";
import type { CssColor, NormalizedFill, NormalizedLine, NormalizedShadow, NormalizedTransform } from "@pdflapas/model";
import { colorToCSS, TRANSPARENT } from "#src/color";
import { UnsupportedFeatureError } from "#src/UnsupportedFeatureError";
import { checkedSize } from "#src/normalize/size";
import { degreesToRadians, millimetersToPixels, pointsToPixels } from "@pdflapas/model";

type ColorValue = string | { color?: string; transparency?: number; alpha?: number; type?: string };

const SCHEME_COLORS: Readonly<Record<string, string>> = Object.freeze({
    tx1: PPTX_DEFAULTS.theme.colors.dark1,
    tx2: PPTX_DEFAULTS.theme.colors.dark2,
    bg1: PPTX_DEFAULTS.theme.colors.light1,
    bg2: PPTX_DEFAULTS.theme.colors.light2,
    accent1: PPTX_DEFAULTS.theme.colors.accent1,
    accent2: PPTX_DEFAULTS.theme.colors.accent2,
    accent3: PPTX_DEFAULTS.theme.colors.accent3,
    accent4: PPTX_DEFAULTS.theme.colors.accent4,
    accent5: PPTX_DEFAULTS.theme.colors.accent5,
    accent6: PPTX_DEFAULTS.theme.colors.accent6,
});

/** Own keys only, so "constructor" or "__proto__" never resolve through the prototype. */
function resolveSchemeColor(color: string): string {
    return Object.hasOwn(SCHEME_COLORS, color) ? (SCHEME_COLORS[color] ?? color) : color;
}

/** Throws on a color that is neither 6-digit hex nor a scheme color (see colorToCSS). */
export function normalizeColor(value?: ColorValue): CssColor {
    if (!value) {
        return TRANSPARENT;
    }

    if (typeof value === "string") {
        return colorToCSS(resolveSchemeColor(value));
    }

    const resolved = value.color ? resolveSchemeColor(value.color) : undefined;

    return colorToCSS({ color: resolved, transparency: value.transparency ?? value.alpha });
}

export function normalizeFill(value?: ColorValue): NormalizedFill {
    if (!value || (typeof value === "object" && value.type === "none")) {
        return { visible: false, color: TRANSPARENT };
    }

    return { visible: true, color: normalizeColor(value) };
}

/** Width and length (px) of a "med" line head, the only size PptxGenJS writes (see PPTX_DEFAULTS.shape.arrowHead). */
function arrowHeadSize(arrow: string, width: number): number {
    const { widthFactor, openArrowWidthFactor, minLineWidthMm } = PPTX_DEFAULTS.shape.arrowHead;

    return (arrow === "arrow" ? openArrowWidthFactor : widthFactor) * Math.max(width, millimetersToPixels(minLineWidthMm));
}

/**
 * Where LibreOffice starts the stroked shaft behind a begin arrowhead (drawinglayer
 * PolygonStrokeArrowPrimitive2D): the head consumes its full length, or half for the centred oval and
 * diamond, less a 1/15 head-width overlap. The dash pattern restarts there, so the phase depends on it.
 * shortcut: also shifts closed outlines that LibreOffice draws without heads; pass the geometry in if that matters.
 */
export function arrowShaftStart(arrow: string, width: number): number {
    if (arrow === "none") {
        return 0;
    }

    const head = arrowHeadSize(arrow, width);

    return head * (arrow === "oval" || arrow === "diamond" ? 0.5 : 1) - head / 15;
}

/**
 * The head outline as closed SVG path data (px), with the line end at the origin and +x pointing out of
 * the line, so renderers only rotate it onto each end. Outlines follow the marker polygons LibreOffice 26.2
 * builds for OOXML heads (oox lineproperties.cxx lclPushMarkerProperties), in percent of the head with the
 * tip at (50, 0): the oval and diamond are centred on the line end, the others put their tip on it. The
 * oval is a true ellipse as in PowerPoint (LibreOffice approximates it with 12 points). An end without a head
 * (`none`) has no outline: the caller leaves it out.
 */
export function arrowHeadPath(arrow: string, width: number): string {
    const size = arrowHeadSize(arrow, width);
    const r = size / 2;

    if (arrow === "oval") {
        return `M ${r} 0 A ${r} ${r} 0 1 0 ${-r} 0 A ${r} ${r} 0 1 0 ${r} 0 Z`;
    }

    // The open arrow is a filled chevron whose arms and stem are one line width thick.
    const half = Math.max((50 * width) / size, 1);
    const percent: Record<string, readonly (readonly [number, number])[]> = {
        triangle: [
            [50, 0],
            [100, 100],
            [0, 100],
        ],
        stealth: [
            [50, 0],
            [100, 100],
            [50, 60],
            [0, 100],
        ],
        diamond: [
            [50, 0],
            [100, 50],
            [50, 100],
            [0, 50],
        ],
        arrow: [
            [50, 0],
            [100, 100 - half * 1.5],
            [100 - half * 1.5, 100],
            [50 + half, 5.5 * half],
            [50 + half, 100],
            [50 - half, 100],
            [50 - half, 5.5 * half],
            [half * 1.5, 100],
            [0, 100 - half * 1.5],
        ],
    };
    const points = percent[arrow];

    if (!points) {
        throw new Error(`Unknown line arrow type '${arrow}': use one of none, arrow, diamond, oval, stealth or triangle`);
    }

    const shift = arrow === "diamond" ? r : 0;

    return `${points.map(([x, y], index) => `${index ? "L" : "M"} ${shift - (y * size) / 100} ${((x - 50) * size) / 100}`).join(" ")} Z`;
}

/**
 * Rotates a dash pattern so its first dash begins `start` px along the path. Renderers take no
 * dash offset, so the split entry is padded with a zero-length dash or gap to keep dash/gap parity.
 */
export function startDashAt(pattern: readonly number[], start: number): number[] {
    let period = 0;

    for (const part of pattern) {
        period += part;
    }

    let offset = (((period - start) % period) + period) % period;

    if (offset === 0) {
        return [...pattern];
    }

    const index = pattern.findIndex(part => (offset -= part) < 0);
    const rest = [-offset, ...pattern.slice(index + 1), ...pattern.slice(0, index), pattern[index]! + offset];

    return index % 2 === 0 ? [...rest, 0] : [0, ...rest];
}

/**
 * An explicit dash array that strokes `pattern` from `start` to `end` px along a path of `length` px and
 * leaves the rest blank. The leading zero-length dash keeps dash/gap parity, as in startDashAt.
 */
export function shaftDashes(pattern: readonly number[], start: number, end: number, length: number): number[] {
    const parts = [0, start];

    for (let at = start, index = 0; at < end; index = (index + 1) % pattern.length) {
        const part = pattern[index] ?? 0;

        parts.push(Math.min(part, end - at));
        at += part;
    }

    // Close on a gap that runs past the end, so the pattern never repeats onto the line.
    if (parts.length % 2) {
        parts.push(length);
    } else {
        parts[parts.length - 1]! += length;
    }

    return parts;
}

/** The path a line strokes: a straight line preset `length` px long, or any other outline. */
export type LineOutline = { kind: "straight"; length: number } | { kind: "other" };

/**
 * `outline` is the path the line strokes. LibreOffice strokes a straight line preset only between
 * its heads (arrowShaftStart at each end), so the stroke neither pokes out of a pointed tip nor fills the
 * stealth notch; the model spells that out as an explicit dash array.
 * shortcut: other open paths keep the full stroke under their heads; trim them by arc length if a fixture needs it.
 * Reads the line object as PptxGenJS 4.0.1 does (addShapeDefinition, addTextDefinition, slideObjectToXml): `width`,
 * `dashType` and the arrow types, never the deprecated `pt`, `size`, `lineDash`, `lineHead` or `lineTail` the
 * type declares on it. The deprecated top-level options are mapped by normalizeShapeLine.
 */
export function normalizeLine(value?: PptxGenJS.ShapeLineProps, outline: LineOutline = { kind: "other" }): NormalizedLine {
    if (!value || value.type === "none") {
        return { visible: false, color: TRANSPARENT, width: 0, style: "solid" };
    }

    const dash = value.dashType ?? "solid";
    const width = pointsToPixels(checkedSize(value.width ?? PPTX_DEFAULTS.shape.line.widthPt, "line.width", PPTX_DEFAULTS.limits.lineWidthPt));
    const beginArrow = value.beginArrowType ?? "none";
    const endArrow = value.endArrowType ?? "none";
    const pattern = width > 0 ? PPTX_DEFAULTS.shape.dashPatterns[dash]?.map(part => part * width) : undefined;
    const trimmed = outline.kind === "straight" && (beginArrow !== "none" || endArrow !== "none");

    // shaftDashes writes one entry per dash and gap along the shaft, so a hair-thin or very long dashed line would
    // allocate without bound; `!(<=)` also refuses the Infinity and NaN a pattern that rounds to zero gives.
    if (trimmed && pattern) {
        let period = 0;

        for (const part of pattern) {
            period += part;
        }

        const entries = (outline.length / period) * pattern.length;

        if (!(entries <= PPTX_DEFAULTS.limits.lineShaftDashEntries)) {
            throw new RangeError(`line: a ${dash} line with an arrow head, ${Math.round(outline.length)} px long and ${width} px wide, needs more than ${PPTX_DEFAULTS.limits.lineShaftDashEntries} dashes; make it wider, shorter or solid`);
        }
    }

    const dashArray = trimmed
        ? shaftDashes(pattern ?? [outline.length, outline.length], arrowShaftStart(beginArrow, width), outline.length - arrowShaftStart(endArrow, width), outline.length)
        : pattern && startDashAt(pattern, arrowShaftStart(beginArrow, width));

    return {
        visible: true,
        color: normalizeColor({ ...value, color: value.color ?? PPTX_DEFAULTS.shape.line.color }),
        width,
        style: dash === "solid" ? "solid" : dash === "sysDot" ? "dotted" : "dashed",
        dashArray: dashArray?.join(" "),
        beginArrow,
        endArrow,
        beginHeadPath: beginArrow === "none" ? undefined : arrowHeadPath(beginArrow, width),
        endHeadPath: endArrow === "none" ? undefined : arrowHeadPath(endArrow, width),
    };
}

/**
 * PptxGenJS 4.0.1 slideObjectToXml writes the inner shadow of a shape or text box as malformed XML (<a:innerShdw>
 * closed by </a:outerShdw>): LibreOffice 26.2 then draws none of the slide, and there is no drawing to match, so the
 * call that adds it fails. A picture's inner shadow is written well-formed and kept.
 */
export function innerShadowError(object: string): UnsupportedFeatureError {
    return new UnsupportedFeatureError(
        `shadow: { type: "inner" } on ${object}`,
        'PptxGenJS 4.0.1 writes an inner shadow on a shape or text box as malformed XML (<a:innerShdw> closed by </a:outerShdw>), and LibreOffice then draws none of the slide. Use shadow type "outer" or remove the shadow; an image keeps its inner shadow.',
        "property",
    );
}

/**
 * `transform` is the object's: the HTML states an offset in the box's own space, under its rotation and flips, as CSS
 * reads `box-shadow` and `drop-shadow()`. An inner shadow turns with its picture (PowerPoint's PDF export, #345); an
 * outer one is cast in slide space, as PptxGenJS writes rotWithShape="0" and LibreOffice 26.2 and PowerPoint draw it
 * (#624), so its offset is turned back by the transform.
 */
export function normalizeShadow(value?: PptxGenJS.ShadowProps, transform: Pick<NormalizedTransform, "rotation" | "flipH" | "flipV"> = { rotation: 0, flipH: false, flipV: false }): NormalizedShadow {
    if (!value || value.type === "none") {
        return { visible: false, color: TRANSPARENT, blur: 0, offsetX: 0, offsetY: 0, inset: false };
    }

    // `||`, not `??`: PptxGenJS replaces 0 with the default too (see the defaults).
    const defaults = PPTX_DEFAULTS.shadow.objectWhenEnabled;
    const angle = degreesToRadians(value.angle || defaults.angle);
    // An inner shadow falls on the side `angle` points to (PowerPoint's PDF export), while CSS's inset shadow is the
    // outside of the box moved by the offset, which darkens the opposite side: so the inset offset points the other way.
    const inner = value.type === "inner";
    const offset = pointsToPixels(checkedSize(value.offset || defaults.offsetPt, "shadow.offset", PPTX_DEFAULTS.limits.shadowOffsetPt)) * (inner ? -1 : 1);
    // CSS blurs a shadow with sigma = blur / 2, so the CSS blur of an inner shadow is twice its measured sigma.
    const blur = pointsToPixels(checkedSize(value.blur || defaults.blurPt, "shadow.blur", PPTX_DEFAULTS.limits.shadowBlurPt)) * (inner ? 2 * PPTX_DEFAULTS.shadow.innerSigmaPerBlur : 1);
    const opacity = value.opacity || defaults.opacity;
    const color = normalizeColor({ color: value.color || defaults.color, transparency: (1 - opacity) * 100 });

    const slideX = Math.cos(angle) * offset;
    const slideY = Math.sin(angle) * offset;
    // CSS maps a box offset onto the slide by flipping it, then rotating it: undo the rotation, then the flips.
    const turn = inner ? 0 : degreesToRadians(-transform.rotation);
    const boxX = slideX * Math.cos(turn) - slideY * Math.sin(turn);
    const boxY = slideX * Math.sin(turn) + slideY * Math.cos(turn);

    return {
        visible: true,
        color,
        blur,
        offsetX: !inner && transform.flipH ? -boxX : boxX,
        offsetY: !inner && transform.flipV ? -boxY : boxY,
        inset: inner,
    };
}
