import type PptxGenJS from "pptxgenjs";
import type { NormalizedObjectStyle } from "@pdflapas/model";
import type { PageSize } from "#src/pageLayouts";
import { normalizeGeometry, normalizeTransform, type PositionOptions, type TransformOptions } from "./geometry";
import { normalizeFill, normalizeLine, normalizeShadow } from "./style";

export type ObjectStyleOptions = {
    objectName?: string;
    fill?: PptxGenJS.ShapeFillProps;
    line?: PptxGenJS.ShapeLineProps;
    shadow?: PptxGenJS.ShadowProps;
} & PositionOptions &
    TransformOptions;

export type NormalizeObjectFlags = {
    fill?: boolean;
    line?: boolean;
    shadow?: boolean;
};

export function normalizeObjectStyle(options: ObjectStyleOptions, defaults: PositionOptions, pageSize: PageSize, objectName: string, flags: NormalizeObjectFlags = {}): NormalizedObjectStyle {
    const transform = normalizeTransform(options);

    return {
        geometry: normalizeGeometry(options, defaults, pageSize),
        transform,
        fill: flags.fill ? normalizeFill(options.fill) : undefined,
        line: flags.line ? normalizeLine(options.line) : undefined,
        shadow: flags.shadow ? normalizeShadow(options.shadow, transform) : undefined,
        objectName: options.objectName ?? objectName,
    };
}
