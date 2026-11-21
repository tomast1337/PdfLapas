import type PptxGenJS from "pptxgenjs";
import type { NormalizedLine } from "@pdflapas/model";
import { PPTX_DEFAULTS } from "#src/defaults";
import { type LineOutline, normalizeLine } from "#src/normalize/style";

/**
 * The outline of addText(). PptxGenJS 4.0.1 (addTextDefinition) reads a string `line` and the deprecated
 * lineSize, lineDash, lineHead and lineTail only for `shape: "line"`; any other text writes <a:ln> from the
 * `line` object alone, without `w` when it has no width.
 */
export function normalizeTextLine(options: PptxGenJS.TextPropsOptions, outline: LineOutline): NormalizedLine {
    // TextPropsOptions declares the same line fields as ShapeProps; normalizeShapeLine reads only those.
    if (options.shape === "line") {
        return normalizeShapeLine(options as unknown as PptxGenJS.ShapeProps, line => normalizeLine(line, outline));
    }

    if (typeof options.line !== "object") {
        return normalizeLine(undefined);
    }

    // `||`: the writer drops a zero width as well.
    return normalizeLine({ ...options.line, width: options.line.width || PPTX_DEFAULTS.text.hairlineWidthPt });
}

/** The outline of a `width` x `height` px shape: straight, with its length, for the straight line presets, whose stroke normalizeLine trims behind its heads. */
export function lineOutline(shapeName: string, width: number, height: number): LineOutline {
    return shapeName === "line" || shapeName === "lineInv" ? { kind: "straight", length: Math.hypot(width, height) } : { kind: "other" };
}

export function normalizeShapeLine(options: PptxGenJS.ShapeProps, normalize: (line?: PptxGenJS.ShapeLineProps) => NormalizedLine): NormalizedLine {
    const hasDeprecatedLine = options.lineSize !== undefined || options.lineDash !== undefined || options.lineHead !== undefined || options.lineTail !== undefined;

    if (!options.line && !hasDeprecatedLine) {
        return normalize(undefined);
    }

    const source = typeof options.line === "string" ? { color: options.line } : (options.line ?? {});

    return normalize({
        ...source,
        width: options.lineSize ?? source.width,
        dashType: options.lineDash ?? source.dashType,
        beginArrowType: options.lineHead ?? source.beginArrowType,
        endArrowType: options.lineTail ?? source.endArrowType,
    });
}
