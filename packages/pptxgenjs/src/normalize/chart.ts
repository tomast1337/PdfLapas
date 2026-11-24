import type PptxGenJS from "pptxgenjs";
import { automaticAxisScale, bar3DAxisScale, type FixedAxisBound } from "#src/chart/axisScale";
import { type ChartTimeUnit, DATE_AXIS_SLOT_LIMIT, dateAxisLayout, formatChartDate, isSupportedChartDateFormat } from "#src/chart/dateAxis";
import { UnsupportedChartError } from "#src/chart/errors";
import { formatChartNumber, isSupportedChartNumberFormat } from "#src/chart/format";
import { PPTX_DEFAULTS } from "#src/defaults";
import { DEFAULT_RENDERING_TARGET, RENDERING_TARGETS, type RenderingTarget, type RenderingTargetProfile } from "#src/target/index";
import type { CssColor, NormalizedChart, NormalizedChartDataLabels, NormalizedChartSeries, NormalizedChartSeriesStyle, NormalizedChartType, NormalizedChartValueAxis } from "@pdflapas/model";
import { checkedSize } from "./size";
import { normalizeColor, normalizeFill } from "./style";

const SUPPORTED_TYPES = new Set<NormalizedChartType>(["area", "bar", "bar3D", "bubble", "doughnut", "line", "mixed", "pie", "radar", "scatter"]);

const MIXED_CHART_TYPES = new Set<PptxGenJS.CHART_NAME>(["area", "bar", "line"]);

export const CATEGORY_AXIS_OPTIONS = [
    "catAxisBaseTimeUnit",
    "catAxisCrossesAt",
    "catAxisHidden",
    "catAxisLabelColor",
    "catAxisLabelFontBold",
    "catAxisLabelFontFace",
    "catAxisLabelFontItalic",
    "catAxisLabelFontSize",
    "catAxisLabelFrequency",
    "catAxisLabelPos",
    "catAxisLabelRotate",
    "catAxisLineColor",
    "catAxisLineShow",
    "catAxisLineSize",
    "catAxisLineStyle",
    "catAxisMajorTickMark",
    "catAxisMajorTimeUnit",
    "catAxisMajorUnit",
    "catAxisMaxVal",
    "catAxisMinorTickMark",
    "catAxisMinorTimeUnit",
    "catAxisMinorUnit",
    "catAxisMinVal",
    "catAxisMultiLevelLabels",
    "catAxisOrientation",
    "catAxisTitle",
    "catAxisTitleColor",
    "catAxisTitleFontFace",
    "catAxisTitleFontSize",
    "catAxisTitleRotate",
    "catGridLine",
    "catLabelFormatCode",
    "showCatAxisTitle",
] as const;

/**
 * The bar3D depth-axis options: the series names on the depth axis, its title and its gridlines (#567, #581). The
 * others draw nothing in LibreOffice 26.2.6 (its PDF export of synthetic PptxGenJS 4.0.1 bar3D charts, four series):
 * serAxisLineShow and serAxisLineColor (it draws no depth-axis line), serAxisLabelPos (PptxGenJS writes `low` on a
 * column chart whatever it says), serLabelFormatCode (the series names are text), serAxisMajorUnit and
 * serAxisMinorUnit (written with serLabelFormatCode only), and the serAxis*TimeUnit options (PptxGenJS checks the
 * option's name, not its value, so it never writes them).
 */
const SERIES_AXIS_OPTIONS = [
    "serAxisBaseTimeUnit",
    "serAxisHidden",
    "serAxisLabelColor",
    "serAxisLabelFontBold",
    "serAxisLabelFontFace",
    "serAxisLabelFontItalic",
    "serAxisLabelFontSize",
    "serAxisLabelFrequency",
    "serAxisLabelPos",
    "serAxisLineColor",
    "serAxisLineShow",
    "serAxisMajorTimeUnit",
    "serAxisMajorUnit",
    "serAxisMinorTimeUnit",
    "serAxisMinorUnit",
    "serAxisOrientation",
    "serGridLine",
    "serLabelFormatCode",
    "serAxisTitle",
    "serAxisTitleColor",
    "serAxisTitleFontFace",
    "serAxisTitleFontSize",
    "serAxisTitleRotate",
    "showSerAxisTitle",
] as const;

/** Options rendered by the initial SVG backend. Everything else is rejected explicitly. */
export const SUPPORTED_CHART_OPTIONS = Object.freeze([
    "x",
    "y",
    "w",
    "h",
    "objectName",
    "altText",
    "chartColors",
    "chartColorsOpacity",
    "fontFace",
    "fontSize",
    "showLegend",
    "legendPos",
    "legendColor",
    "legendFontFace",
    "legendFontSize",
    "showTitle",
    "title",
    "titleBold",
    "titleColor",
    "titleFontFace",
    "titleFontSize",
    "titlePos",
    "titleRotate",
    "bar3DShape",
    "barDir",
    "barGrouping",
    "barGapDepthPct",
    "barGapWidthPct",
    "barOverlapPct",
    "invertedColors",
    "dataBorder",
    "dataLabelBkgrdColors",
    "showDataTable",
    "showDataTableHorzBorder",
    "showDataTableKeys",
    "showDataTableOutline",
    "showDataTableVertBorder",
    "dataTableFontSize",
    "dataTableFormatCode",
    "v3DPerspective",
    "v3DRAngAx",
    "v3DRotX",
    "v3DRotY",
    "showValAxisTitle",
    "valAxes",
    // Each entry's category-axis options over the chart's (#594); a second entry is a secondary category axis.
    "catAxes",
    "valAxisCrossesAt",
    "valAxisDisplayUnit",
    "valAxisDisplayUnitLabel",
    "valAxisHidden",
    "valAxisMaxVal",
    "valAxisMinVal",
    "valAxisTitle",
    "valAxisLabelColor",
    "valAxisLabelFontBold",
    "valAxisLabelFontFace",
    "valAxisLabelFontItalic",
    "valAxisLabelFontSize",
    "valAxisLabelFormatCode",
    "valAxisLabelPos",
    "valAxisLabelRotate",
    "valAxisLineColor",
    "valAxisLineShow",
    "valAxisLineSize",
    "valAxisLineStyle",
    "valAxisLogScaleBase",
    "valAxisMajorTickMark",
    "valAxisMajorUnit",
    "valAxisMinorTickMark",
    "valAxisOrientation",
    "valAxisTitleColor",
    "valAxisTitleFontFace",
    "valAxisTitleFontSize",
    "valAxisTitleRotate",
    "valGridLine",
    "holeSize",
    "firstSliceAng",
    "radarStyle",
    "lineSmooth",
    "lineDataSymbol",
    "lineDataSymbolLineColor",
    "lineDataSymbolLineSize",
    "lineDataSymbolSize",
    "lineSize",
    "showValue",
    "showLabel",
    "showPercent",
    "showSerName",
    // PptxGenJS 4.0.1 writes `<c:showLeaderLines>` from it on every chart but pie and doughnut (always 1 there), and
    // Office draws a leader line only to a label moved away from its point, which PptxGenJS never writes.
    "showLeaderLines",
    // Not a PptxGenJS 4.0.1 option: every writer prints `<c:showLegendKey val="0"/>` whatever it says.
    "showLegendKey",
    "dataLabelPosition",
    "dataLabelFormatCode",
    "dataLabelColor",
    "dataLabelFontBold",
    "dataLabelFontFace",
    "dataLabelFontItalic",
    "dataLabelFontSize",
    ...CATEGORY_AXIS_OPTIONS,
    ...SERIES_AXIS_OPTIONS,
    "layout",
    "plotArea",
    // Deprecated v3.11 aliases PptxGenJS 4.0.1 still reads: `plotArea.fill.color = fill`, `plotArea.border = border`.
    "fill",
    "border",
] as const);

// The chart types on which a target may draw no plot-area fill or border (drawsPlotAreaOnEveryType, #479, #565);
// `layout` places each of them.
const PLOT_AREA_UNDRAWN_TYPES = new Set<NormalizedChartType>(["bar3D", "doughnut", "pie", "radar"]);
const SUPPORTED_PLOT_AREA_OPTIONS: Readonly<Record<string, ReadonlySet<string>>> = Object.freeze({
    // genXmlColorSelection reads type, color, transparency and the deprecated alpha.
    fill: new Set(["alpha", "color", "transparency", "type"]),
    // PptxGenJS writes pt and color and ignores type: the border is always a solid line.
    border: new Set(["color", "pt", "type"]),
});
const LAYOUT_KEYS = ["x", "y", "w", "h"] as const;

const SUPPORTED_CHART_OPTION_SET = new Set<string>(SUPPORTED_CHART_OPTIONS);
// A combo chart writes each chart type with `{ ...options, ...type.options }`, so a type's own label options apply to
// its series.
const MIXED_LABEL_OPTIONS = [
    "dataLabelColor",
    "dataLabelFontBold",
    "dataLabelFontFace",
    "dataLabelFontItalic",
    "dataLabelFontSize",
    "dataLabelFormatCode",
    "dataLabelPosition",
    "showLabel",
    "showLeaderLines",
    "showLegendKey",
    "showPercent",
    "showSerName",
    "showValue",
] as const;
// A combo chart type's own fill, outline, line and marker options, which PptxGenJS writes on that type's series only
// (#590, #594): each series of the type takes them as its style.
const MIXED_STYLE_OPTIONS = ["chartColorsOpacity", "dataBorder", "lineDataSymbol", "lineDataSymbolLineColor", "lineDataSymbolLineSize", "lineDataSymbolSize", "lineSize", "lineSmooth"] as const;
// A bar type's gap and overlap, drawn chart-wide with one value per kind (#590).
const MIXED_BAR_OPTIONS = ["barGapWidthPct", "barOverlapPct"] as const;
const SUPPORTED_MIXED_SERIES_OPTIONS = new Set<string>(["barGrouping", "chartColors", "invertedColors", "secondaryCatAxis", "secondaryValAxis", ...MIXED_LABEL_OPTIONS, ...MIXED_STYLE_OPTIONS, ...MIXED_BAR_OPTIONS]);
// The options a catAxes entry may set (#594).
const CATEGORY_AXIS_OPTION_SET = new Set<string>(CATEGORY_AXIS_OPTIONS);
// The options catAxes[1] may set: the secondary category axis's labels and title, as PowerPoint and LibreOffice 26.2.6
// draw them (#603).
const SECONDARY_CATEGORY_AXIS_OPTIONS = new Set<string>([
    "catAxisHidden",
    "catAxisLabelColor",
    "catAxisLabelFontBold",
    "catAxisLabelFontFace",
    "catAxisLabelFontItalic",
    "catAxisLabelFontSize",
    "catAxisLabelPos",
    "catAxisOrientation",
    "catAxisTitle",
    "catAxisTitleColor",
    "catAxisTitleFontFace",
    "catAxisTitleFontSize",
    "catAxisTitleRotate",
    "showCatAxisTitle",
]);
const SUPPORTED_VALUE_AXIS_OPTIONS = new Set([
    "showValAxisTitle",
    "valAxisDisplayUnit",
    "valAxisDisplayUnitLabel",
    "valAxisHidden",
    "valAxisMaxVal",
    "valAxisMinVal",
    "valAxisTitle",
    "valAxisLabelColor",
    "valAxisLabelFontBold",
    "valAxisLabelFontFace",
    "valAxisLabelFontItalic",
    "valAxisLabelFontSize",
    "valAxisLabelFormatCode",
    "valAxisLabelPos",
    "valAxisLabelRotate",
    "valAxisLineColor",
    "valAxisLineShow",
    "valAxisLineSize",
    "valAxisLineStyle",
    "valAxisLogScaleBase",
    "valAxisMajorTickMark",
    "valAxisMajorUnit",
    "valAxisMinorTickMark",
    "valAxisOrientation",
    "valAxisTitleColor",
    "valAxisTitleFontFace",
    "valAxisTitleFontSize",
    "valAxisTitleRotate",
    "valGridLine",
]);
// A date axis's time line (#580): drawn on line, column, bar, area and combo charts, not measured on radar or bar3D ones.
const DATE_AXIS_OPTIONS = ["catAxisBaseTimeUnit", "catAxisMajorTimeUnit", "catAxisMinorTimeUnit", "catAxisMajorUnit", "catAxisMinorUnit", "catAxisMinVal", "catAxisMaxVal"] as const;
const DATE_AXIS_TYPES = new Set<NormalizedChartType>(["area", "bar", "line", "mixed"]);
const TIME_UNITS = new Set(["days", "months", "years"]);
// The category-axis options a chart type does not draw (#566, #582). A radar chart's category axis is its spokes:
// LibreOffice 26.2.6 draws its labels in catAxisLabel* (none for catAxisHidden or catAxisLabelPos none), runs them
// counterclockwise for catAxisOrientation maxMin and numbers the value axis on the spoke a numeric catAxisCrossesAt
// names, and draws neither a catAxisTitle nor the catGridLine (#479); the line, tick, rotation, skip and format options
// are not measured there. A bar3D chart draws its crossings as LibreOffice 26.2.6 and PowerPoint do (#591): the bars grow
// from valAxisCrossesAt and the value labels stand at catAxisCrossesAt.
const CATEGORY_AXIS_UNDRAWN_OPTIONS: Readonly<Partial<Record<NormalizedChartType, ReadonlySet<string>>>> = Object.freeze({
    radar: new Set<string>([
        ...DATE_AXIS_OPTIONS,
        "catAxisLabelFrequency",
        "catLabelFormatCode",
        "valAxisCrossesAt",
        "catGridLine",
        "catAxisLabelRotate",
        "catAxisLineColor",
        "catAxisLineShow",
        "catAxisLineSize",
        "catAxisLineStyle",
        "catAxisMajorTickMark",
        "catAxisMinorTickMark",
    ]),
    bar3D: new Set<string>(DATE_AXIS_OPTIONS),
});
// PptxGenJS 4.0.1 writes `<c:dTable>` on every chart type with showDataTable. LibreOffice 26.2.6 draws none on a pie,
// doughnut or radar chart (#583); it draws a scatter or bubble chart's as one malformed row (drawsScatterDataTable).
// LibreOffice and PowerPoint draw a bar3D chart's under the box (#592).
const DATA_TABLE_UNDRAWN_TYPES = new Set<NormalizedChartType>(["doughnut", "pie", "radar"]);
const XY_TYPES = new Set<NormalizedChartType>(["scatter", "bubble"]);
const DATA_BORDER_TYPES = new Set<NormalizedChartType>(["area", "bar", "bar3D", "bubble", "doughnut", "mixed", "pie"]);
// PptxGenJS types its tick marks "none", "inside", "outside" and "cross" and writes the value into `<c:majorTickMark>` and
// `<c:minorTickMark>` as given, so the OOXML names "in" and "out" work too; LibreOffice draws "inside" and "outside" as
// them (PowerPoint likely cannot open a file with them, #630).
const TICK_MARKS: Readonly<Record<string, NormalizedChartValueAxis["majorTickMark"]>> = { none: "none", inside: "inside", outside: "outside", cross: "cross", in: "inside", out: "outside" };
// bar3D draws its value axis on a wall, its line and tick marks where the target draws them (drawsBar3DValueAxisLine);
// it draws no logarithmic or reversed 3D axis or horizontal 3D bars yet.
// A radar chart's value axis is its first spoke: LibreOffice draws its labels and rings, the line style on every spoke,
// and no title (#561).
const VALUE_AXIS_UNDRAWN_OPTIONS: Readonly<Partial<Record<NormalizedChartType, ReadonlySet<string>>>> = Object.freeze({
    bar3D: new Set(["valAxisLogScaleBase", "valAxisOrientation"]),
    radar: new Set(["valAxisLogScaleBase", "valAxisOrientation", "valAxisLabelRotate"]),
});
const SUPPORTED_GRID_LINE_OPTIONS = new Set(["cap", "color", "size", "style"]);
const BAR_3D_SHAPES = new Set(["box", "cylinder", "cone", "coneToMax", "pyramid", "pyramidToMax"]);
// The dataLabelPosition values PptxGenJS 4.0.1 keeps; its option cleanup deletes any other (every one on an area, radar,
// bar3D or doughnut chart and, a quirk, `outEnd` on a clustered bar), so the chart type's default position applies. A
// combo chart's skips the cleanup, so each chart type keeps the positions valid for it.
const LINE_LABEL_POSITIONS = new Set(["b", "ctr", "l", "r", "t"]);
const CLUSTERED_BAR_LABEL_POSITIONS = new Set(["ctr", "inBase", "inEnd"]);
const STACKED_BAR_LABEL_POSITIONS = new Set(["ctr", "inBase", "inEnd", "outEnd"]);
const PIE_LABEL_POSITIONS = new Set(["bestFit", "ctr", "inEnd", "outEnd"]);
const NO_LABEL_POSITIONS = new Set<string>();

function normalizeSeries(type: NormalizedChartType, data: readonly PptxGenJS.OptsChartData[]): NormalizedChartSeries[] {
    if (data.length === 0) {
        throw new TypeError("Chart data must contain at least one series");
    }

    const normalized: NormalizedChartSeries[] = [];

    for (const [index, series] of data.entries()) {
        const rawLabels = series.labels ?? [];

        if (rawLabels.some(Array.isArray)) {
            throw new UnsupportedChartError({
                reason: "chart-option",
                chartTypes: [type],
                unsupportedOptions: ["data.labels[multi-level]"],
            });
        }

        const labels = (rawLabels as string[]).map(label => String(label));
        const values: number[] = [];

        for (const [valueIndex, value] of (series.values ?? []).entries()) {
            if (typeof value !== "number" || !Number.isFinite(value)) {
                throw new TypeError(`Chart series ${index} value ${valueIndex} must be a finite number`);
            }

            values.push(value);
        }

        let sizes: number[] | undefined;

        if (series.sizes) {
            sizes = [];

            for (const [valueIndex, value] of series.sizes.entries()) {
                if (typeof value !== "number" || !Number.isFinite(value) || value < 0) {
                    throw new TypeError(`Chart series ${index} bubble size ${valueIndex} must be a non-negative finite number`);
                }

                sizes.push(value);
            }
        }

        normalized.push({
            type,
            name: series.name ?? `Series ${index + 1}`,
            labels,
            values,
            sizes,
            colorIndex: index,
            varyColors: false,
            valueAxisIndex: 0,
        });
    }

    return normalized;
}

/** The lowest and highest values a value axis shows: each value, or each stack's totals or shares (0 to 0 with none). */
export function valueAxisDataRange(type: NormalizedChartType, series: readonly NormalizedChartSeries[]): [number, number] {
    const values: number[] = [];
    // The stacked series by type: each type stacks apart from the others.
    const stacks = new Map<NormalizedChartType, NormalizedChartSeries[]>();

    for (const [index, item] of series.entries()) {
        // A scatter or bubble chart's first series holds the X values.
        if ((type === "scatter" || type === "bubble") && index === 0) {
            continue;
        }

        if (!item.stack) {
            values.push(...item.values);
            continue;
        }

        const stack = stacks.get(item.type);

        if (stack) {
            stack.push(item);
        } else {
            stacks.set(item.type, [item]);
        }
    }

    // Each stack's positive and negative totals per category, or their shares of the category's absolute total (in
    // percent, or as fractions of 1).
    for (const stack of stacks.values()) {
        const share = stack[0]?.stack === "percentStacked" ? 100 : stack[0]?.stack === "shareStacked" ? 1 : undefined;
        const categoryCount = Math.max(0, ...stack.map(item => item.values.length));

        for (let index = 0; index < categoryCount; index++) {
            let positive = 0;
            let negative = 0;

            for (const item of stack) {
                const value = item.values[index] ?? 0;

                if (value >= 0) {
                    positive += value;
                } else {
                    negative += value;
                }
            }

            const total = positive - negative;

            if (share === undefined) {
                values.push(positive, negative);
            } else if (total === 0) {
                values.push(0, 0);
            } else {
                values.push((positive / total) * share, (negative / total) * share);
            }
        }
    }

    return values.length > 0 ? [Math.min(...values), Math.max(...values)] : [0, 0];
}

function positiveNumber(value: number | undefined, name: string, fallback: number): number {
    const normalized = value ?? fallback;

    if (!Number.isFinite(normalized) || normalized <= 0) {
        throw new TypeError(`${name} must be a positive finite number`);
    }

    return normalized;
}

function tickMark(value: unknown, name: string, fallback: NormalizedChartValueAxis["majorTickMark"]): NormalizedChartValueAxis["majorTickMark"] {
    if (value === undefined) {
        return fallback;
    }

    const mark = Object.hasOwn(TICK_MARKS, String(value)) ? TICK_MARKS[String(value)] : undefined;

    if (mark === undefined) {
        throw new TypeError(`${name} must be "none", "inside" (or "in"), "outside" (or "out") or "cross"; got ${JSON.stringify(value)}`);
    }

    return mark;
}

function finiteNumber(value: number | undefined, name: string, fallback: number): number {
    const normalized = value ?? fallback;

    if (!Number.isFinite(normalized)) {
        throw new TypeError(`${name} must be a finite number`);
    }

    return normalized;
}

/**
 * How a value axis is scaled: LibreOffice's automatic 2D scale, its bar3D scale, or a radar scale (the target's
 * `radarScale`: the bar3D or the 2D rounding, in at most `maxIncrements` steps).
 */
type ValueAxisScaleKind = { kind: "linear" } | { kind: "bar3D" } | { kind: "radar"; rounding: "bar3D" | "linear"; maxIncrements: number };

function normalizeValueAxis(
    axisOptions: PptxGenJS.IChartPropsAxisVal | undefined,
    dataRange: [number, number],
    scaleKind: ValueAxisScaleKind,
    path: string,
    chartTypes: PptxGenJS.CHART_NAME[],
    defaultGridLine: PptxGenJS.OptsChartGridLine,
    readsDisplayUnits: boolean,
): NormalizedChart["valueAxes"][number] {
    const defaults = PPTX_DEFAULTS.chart.valueAxis;
    // Check the axis's unit, bounds, scale, tick marks and label format. An automatic bound scales from a fixed one (#556).
    const majorUnit = axisOptions?.valAxisMajorUnit;

    if (majorUnit !== undefined && (!Number.isFinite(majorUnit) || majorUnit <= 0)) {
        throw new TypeError("valAxisMajorUnit must be a positive finite number");
    }

    const fixedMinimum = axisOptions?.valAxisMinVal;
    const fixedMaximum = axisOptions?.valAxisMaxVal;
    const fixed: FixedAxisBound = fixedMinimum !== undefined ? { kind: "minimum", value: fixedMinimum } : fixedMaximum !== undefined ? { kind: "maximum", value: fixedMaximum } : { kind: "none" };
    // A bar3D axis takes LibreOffice's 3D scale over its data with 0 included, and renderBar3DSvg takes a longer step
    // where the drawn axis is short; a radar axis takes the target's radarScale over the same range.
    const bar3D = scaleKind.kind === "bar3D";
    const bar3DLow = Math.min(0, dataRange[0]);
    const bar3DHigh = Math.max(0, dataRange[1]);
    const fixedScale = { minimum: fixedMinimum, maximum: fixedMaximum, majorUnit };
    const scale =
        scaleKind.kind === "bar3D"
            ? bar3DAxisScale(bar3DLow, bar3DHigh, fixedScale, PPTX_DEFAULTS.chart.automaticScale.maxIncrements)
            : scaleKind.kind === "radar"
              ? scaleKind.rounding === "bar3D"
                  ? bar3DAxisScale(bar3DLow, bar3DHigh, fixedScale, scaleKind.maxIncrements)
                  : automaticAxisScale(bar3DLow, bar3DHigh, fixed, true, majorUnit !== undefined ? { kind: "fixed", interval: majorUnit } : { kind: "automatic", maxIncrements: scaleKind.maxIncrements })
              : majorUnit !== undefined
                ? automaticAxisScale(dataRange[0], dataRange[1], fixed, true, { kind: "fixed", interval: majorUnit })
                : automaticAxisScale(dataRange[0], dataRange[1], fixed, true);
    const minimum = fixedMinimum ?? scale.minimum;
    const maximum = fixedMaximum ?? scale.maximum;

    if (!Number.isFinite(minimum) || !Number.isFinite(maximum) || minimum >= maximum) {
        throw new TypeError("Chart value-axis minimum must be less than its maximum");
    }

    const logScaleBase = axisOptions?.valAxisLogScaleBase;

    if (logScaleBase !== undefined && (!Number.isFinite(logScaleBase) || logScaleBase < 2 || logScaleBase > 99)) {
        throw new RangeError("valAxisLogScaleBase must be between 2 and 99");
    }

    if (logScaleBase !== undefined && minimum <= 0) {
        throw new TypeError("A logarithmic value axis requires a positive minimum");
    }

    const labelFormatCode = axisOptions?.valAxisLabelFormatCode ?? PPTX_DEFAULTS.chart.valueAxis.labelFormatCode;

    if (!isSupportedChartNumberFormat(labelFormatCode)) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["valAxisLabelFormatCode[format]"] });
    }

    // PptxGenJS replaces a missing valGridLine whole (`valGridLine || default`), so a set one keeps no default field.
    const gridLine = axisOptions?.valGridLine ?? defaultGridLine;

    // PptxGenJS writes `<c:orientation val="${valAxisOrientation || 'minMax'}"/>`; its types name only "minMax", but
    // OOXML's "maxMin" reverses the axis, so the option is read as a string.
    const orientation: string | undefined = axisOptions?.valAxisOrientation;

    if (orientation !== undefined && orientation !== "minMax" && orientation !== "maxMin") {
        throw new TypeError(`${path}valAxisOrientation must be "minMax" or "maxMin"; got ${JSON.stringify(orientation)}`);
    }

    // `<c:dispUnits><c:builtInUnit val="${valAxisDisplayUnit}"/>`, with `<c:dispUnitsLbl/>` for valAxisDisplayUnitLabel.
    const displayUnitName: string | undefined = axisOptions?.valAxisDisplayUnit;
    const displayUnits = PPTX_DEFAULTS.chart.displayUnits;
    let displayUnit: NormalizedChartValueAxis["displayUnit"];

    if (displayUnitName !== undefined) {
        if (!Object.hasOwn(displayUnits, displayUnitName)) {
            throw new TypeError(`${path}valAxisDisplayUnit must be one of ${Object.keys(displayUnits).join(", ")}; got ${JSON.stringify(displayUnitName)}`);
        }

        // Checked just above: the name is one of the table's own keys.
        const unit = displayUnits[displayUnitName as keyof typeof displayUnits];

        if (readsDisplayUnits && axisOptions?.valAxisDisplayUnitLabel && axisOptions.showValAxisTitle) {
            throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: [`${path}valAxisDisplayUnitLabel[showValAxisTitle]`] });
        }

        if (readsDisplayUnits) {
            displayUnit = { divisor: unit.divisor, ...(axisOptions?.valAxisDisplayUnitLabel ? { label: unit.label } : {}) };
        }
    }

    return {
        minimum,
        maximum,
        ...(displayUnit ? { displayUnit } : {}),
        // A radar axis has no step of its own to fall back on (automaticAxisInterval would take ten), so it keeps the scale's.
        majorUnit: scaleKind.kind === "radar" ? scale.interval : majorUnit,
        ...(orientation === "maxMin" ? { reversed: true } : {}),
        ...(bar3D ? { automaticScale: { low: bar3DLow, high: bar3DHigh, fixedMinimum: fixedMinimum !== undefined, fixedMaximum: fixedMaximum !== undefined } } : {}),
        // A 2D axis's automatic step depends on the room its labels have too (#585), which the renderer knows.
        ...(scaleKind.kind === "linear" && majorUnit === undefined && logScaleBase === undefined ? { automaticScale: { low: dataRange[0], high: dataRange[1], fixedMinimum: fixedMinimum !== undefined, fixedMaximum: fixedMaximum !== undefined } } : {}),
        logScaleBase,
        labelFormatCode,
        hidden: axisOptions?.valAxisHidden ?? defaults.hidden,
        showTitle: axisOptions?.showValAxisTitle ?? false,
        title: axisOptions?.valAxisTitle ?? defaults.title,
        titleColor: normalizeColor(axisOptions?.valAxisTitleColor ?? defaults.titleColor),
        titleFontFace: axisOptions?.valAxisTitleFontFace ?? PPTX_DEFAULTS.chart.fontFace,
        titleFontSize: positiveNumber(axisOptions?.valAxisTitleFontSize, "valAxisTitleFontSize", defaults.titleFontSizePt),
        titleRotate: axisOptions?.valAxisTitleRotate === undefined ? defaults.titleRotate : finiteNumber(axisOptions.valAxisTitleRotate, "valAxisTitleRotate", 0),
        labelColor: normalizeColor(axisOptions?.valAxisLabelColor ?? defaults.labelColor),
        labelFontFace: axisOptions?.valAxisLabelFontFace ?? PPTX_DEFAULTS.chart.fontFace,
        labelFontSize: positiveNumber(axisOptions?.valAxisLabelFontSize, "valAxisLabelFontSize", defaults.labelFontSizePt),
        labelBold: axisOptions?.valAxisLabelFontBold ?? defaults.labelBold,
        labelItalic: axisOptions?.valAxisLabelFontItalic ?? defaults.labelItalic,
        labelRotate: finiteNumber(axisOptions?.valAxisLabelRotate, "valAxisLabelRotate", defaults.labelRotate),
        labelPosition: axisOptions?.valAxisLabelPos ?? defaults.labelPosition,
        lineColor: normalizeColor(axisOptions?.valAxisLineColor ?? defaults.lineColor),
        lineWidth: positiveNumber(axisOptions?.valAxisLineSize, "valAxisLineSize", defaults.lineWidthPt),
        lineStyle: axisOptions?.valAxisLineStyle ?? defaults.lineStyle,
        lineVisible: axisOptions?.valAxisLineShow ?? defaults.lineVisible,
        majorTickMark: tickMark(axisOptions?.valAxisMajorTickMark, "valAxisMajorTickMark", defaults.majorTickMark),
        minorTickMark: tickMark(axisOptions?.valAxisMinorTickMark, "valAxisMinorTickMark", defaults.minorTickMark),
        gridLineColor: normalizeColor(gridLine.color ?? PPTX_DEFAULTS.chart.gridLine.color),
        gridLineWidth: positiveNumber(gridLine.size, "valGridLine.size", PPTX_DEFAULTS.chart.gridLine.widthPt),
        gridLineStyle: gridLine.style ?? PPTX_DEFAULTS.chart.gridLine.style,
        gridLineCap: gridLine.cap ?? PPTX_DEFAULTS.chart.gridLine.cap,
    };
}

/**
 * A series' fill opacity, outline, line and markers from the chart's options, or a combo chart type's over them.
 * PptxGenJS writes `<a:alpha>` only for a truthy chartColorsOpacity (in percent), so 0 is opaque;
 * `lineSymbolLineSize` is in points, already converted by the caller.
 */
function normalizeSeriesStyle(options: PptxGenJS.IChartOpts, dataBorder: NormalizedChartSeriesStyle["dataBorder"], lineSymbolLineSize: number, lineSmooth: boolean): NormalizedChartSeriesStyle {
    return {
        colorOpacity: Math.max(0, Math.min(1, (options.chartColorsOpacity || 100) / 100)),
        ...(dataBorder ? { dataBorder } : {}),
        lineSmooth,
        lineSymbol: options.lineDataSymbol ?? PPTX_DEFAULTS.chart.lineDataSymbol,
        lineSymbolSize: checkedSize(options.lineDataSymbolSize ?? PPTX_DEFAULTS.chart.lineDataSymbolSize, "lineDataSymbolSize", PPTX_DEFAULTS.limits.chartMarkerSizePt),
        // PptxGenJS: `<a:ln w="${lineDataSymbolLineSize}"><a:solidFill>${lineDataSymbolLineColor || seriesColor}`.
        ...(options.lineDataSymbolLineColor ? { lineSymbolLineColor: normalizeColor(options.lineDataSymbolLineColor) } : {}),
        lineSymbolLineSize: checkedSize(lineSymbolLineSize, "lineDataSymbolLineSize", PPTX_DEFAULTS.limits.lineWidthPt),
        lineSize: checkedSize(options.lineSize ?? PPTX_DEFAULTS.chart.lineSizePt, "lineSize", PPTX_DEFAULTS.limits.lineWidthPt),
    };
}

/**
 * The data labels PptxGenJS 4.0.1 writes for a chart type (a combo chart's: each of its types) from the label options,
 * as LibreOffice draws them. It writes showCatName and showPercent only on pie and doughnut points and a custom label
 * per point for showLabel on scatter, so on other charts showLabel and showPercent draw nothing, as in PowerPoint; and
 * showSerName on every chart but pie and doughnut.
 */
function normalizeDataLabels(type: NormalizedChartType, options: PptxGenJS.IChartOpts, grouping: NormalizedChart["grouping"], chartTypes: PptxGenJS.CHART_NAME[], profile: RenderingTargetProfile): NormalizedChartDataLabels {
    const pieLike = type === "pie" || type === "doughnut";
    const stackedBars = grouping === "stacked" || grouping === "percentStacked";
    const defaultPositions = PPTX_DEFAULTS.chart.dataLabelPosition;
    let position: NormalizedChartDataLabels["position"];
    let keptPositions: ReadonlySet<string> = NO_LABEL_POSITIONS;

    switch (type) {
        case "line":
            position = defaultPositions.line;
            keptPositions = LINE_LABEL_POSITIONS;
            break;
        case "scatter":
        case "bubble":
            position = defaultPositions.scatter;
            keptPositions = LINE_LABEL_POSITIONS;
            break;
        case "pie":
            position = profile.pieUnsetLabelPosition;
            keptPositions = PIE_LABEL_POSITIONS;
            break;
        case "doughnut":
            position = defaultPositions.doughnut;
            break;
        case "area":
            position = defaultPositions.area;
            break;
        case "radar":
            position = defaultPositions.radar;
            break;
        case "bar3D":
            position = defaultPositions.bar3D;
            break;
        case "bar":
        case "mixed":
            position = stackedBars ? defaultPositions.stacked : defaultPositions.clustered;
            keptPositions = stackedBars ? STACKED_BAR_LABEL_POSITIONS : CLUSTERED_BAR_LABEL_POSITIONS;
            break;
    }

    if (options.dataLabelPosition !== undefined && keptPositions.has(options.dataLabelPosition)) {
        position = options.dataLabelPosition;
    }

    // PptxGenJS: `dataLabelFormatCode || (pie or doughnut: showPercent ? '0%' : 'General'; scatter: 'General') || '#,##0'`.
    // On a pie LibreOffice applies it to the percentage when showPercent is set and draws the value in General (renders
    // of PptxGenJS pies: `0.0` with showPercent draws 0.5; showValue with showPercent draws 50.25 and 50%).
    const pieFormatCode = options.showPercent ? PPTX_DEFAULTS.chart.piePercentFormatCode : PPTX_DEFAULTS.chart.pieDataLabelFormatCode;
    const formatCode = options.dataLabelFormatCode || (pieLike ? pieFormatCode : type === "scatter" ? PPTX_DEFAULTS.chart.scatterDataLabelFormatCode : PPTX_DEFAULTS.chart.dataLabelFormatCode);

    if (!isSupportedChartNumberFormat(formatCode)) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["dataLabelFormatCode[format]"] });
    }

    // A pie or doughnut prints the series name on its own line above the slice's other fields, as LibreOffice 26.2.6
    // and PowerPoint print "Share" / "40" (#565).
    const showSeriesName = options.showSerName ?? PPTX_DEFAULTS.chart.showSeriesName;

    return {
        showValue: options.showValue ?? PPTX_DEFAULTS.chart.showValue,
        showCategory: (pieLike || type === "scatter") && (options.showLabel ?? PPTX_DEFAULTS.chart.showLabel),
        ...(showSeriesName ? { showSeriesName, fieldSeparator: profile.dataLabelFieldSeparator } : {}),
        showPercent: pieLike && (options.showPercent ?? PPTX_DEFAULTS.chart.showPercent),
        position,
        valueFormat: pieLike && options.showPercent ? PPTX_DEFAULTS.chart.pieDataLabelFormatCode : formatCode,
        percentFormat: formatCode,
        color: normalizeColor(options.dataLabelColor || PPTX_DEFAULTS.chart.dataLabelColor),
        fontFace: options.dataLabelFontFace || PPTX_DEFAULTS.chart.fontFace,
        fontSize: checkedSize(options.dataLabelFontSize || PPTX_DEFAULTS.chart.dataLabelFontSizePt, "dataLabelFontSize", PPTX_DEFAULTS.limits.fontSizePt),
        bold: options.dataLabelFontBold ?? false,
        italic: options.dataLabelFontItalic ?? false,
        // PptxGenJS writes the `<c:dLbls>` fill in the series color on bar, line and area series only.
        ...(options.dataLabelBkgrdColors && (type === "bar" || type === "line" || type === "area" || type === "mixed") ? { seriesColorBackground: true } : {}),
    };
}

type ChartInputType = PptxGenJS.CHART_NAME | PptxGenJS.IChartMulti[];

export function normalizeChart(type: ChartInputType, data: readonly PptxGenJS.OptsChartData[], requestedOptions: PptxGenJS.IChartOpts = {}, target: RenderingTarget = DEFAULT_RENDERING_TARGET): NormalizedChart {
    const profile = RENDERING_TARGETS[target];
    const mixed = Array.isArray(type);
    const chartTypes = mixed ? type.map(chart => chart.type) : [type];
    // A combo chart's bar options as its bar type sets them (#590).
    const typeOptions: Record<string, unknown> = {};
    const typeOptionSetBy = new Map<string, number>();
    const unsupportedTypeOptions: string[] = [];

    for (const [index, chart] of (mixed ? type : []).entries()) {
        for (const key of chart.type === "bar" ? MIXED_BAR_OPTIONS : []) {
            const value: unknown = chart.options?.[key as keyof PptxGenJS.IChartOpts];

            if (value === undefined) {
                continue;
            }

            const earlier = typeOptionSetBy.get(key);

            if (earlier !== undefined && typeOptions[key] !== value) {
                unsupportedTypeOptions.push(`series[${index}].options.${key}[differs from series[${earlier}]]`);
            }

            typeOptions[key] = value;
            typeOptionSetBy.set(key, index);
        }
    }

    if (unsupportedTypeOptions.length) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: unsupportedTypeOptions });
    }

    // PptxGenJS 4.0.1 writes the primary category axis from `Object.assign({}, options, catAxes[0])` (#594). Every
    // typeOptions key comes from MIXED_BAR_OPTIONS, read from the type's own IChartOpts.
    const options: PptxGenJS.IChartOpts = { ...requestedOptions, ...requestedOptions.catAxes?.[0], ...(typeOptions as PptxGenJS.IChartOpts) };

    // Refuse a chart type, a combination or an option the renderers do not draw.
    if (mixed && (type.length === 0 || type.some(chart => !MIXED_CHART_TYPES.has(chart.type)))) {
        throw new UnsupportedChartError({ reason: "chart-combination", chartTypes });
    }

    if (!mixed && !SUPPORTED_TYPES.has(type as NormalizedChartType)) {
        throw new UnsupportedChartError({ reason: "chart-type", chartTypes: [String(type)] });
    }

    if (mixed) {
        const unsupported: string[] = [];

        for (const [index, chart] of type.entries()) {
            for (const key of Object.keys(chart.options ?? {})) {
                const value: unknown = chart.options[key as keyof PptxGenJS.IChartOpts];

                if (value !== undefined && !SUPPORTED_MIXED_SERIES_OPTIONS.has(key)) {
                    unsupported.push(`series[${index}].options.${key}`);
                }

                // A secondary category axis is written only from catAxes[1], and crosses the secondary value axis (#594).
                if (key === "secondaryCatAxis" && value === true && (!chart.options.secondaryValAxis || (options.catAxes?.length ?? 0) < 2)) {
                    unsupported.push(`series[${index}].options.secondaryCatAxis[needs secondaryValAxis and catAxes[1]]`);
                }

                // The secondary category labels are measured under a column chart's labels only (#603).
                if (key === "secondaryCatAxis" && value === true && options.barDir === "bar") {
                    unsupported.push(`series[${index}].options.secondaryCatAxis[barDir bar]`);
                }

                // A data table takes the category labels' band; the secondary row beside it is not measured (#603).
                if (key === "secondaryCatAxis" && value === true && options.showDataTable) {
                    unsupported.push(`series[${index}].options.secondaryCatAxis[showDataTable]`);
                }

                // PptxGenJS writes a type's own dataBorder as it is, `w="${valToPts(pt)}"` in `color`, unchecked.
                if (key === "dataBorder" && value !== undefined) {
                    const border = chart.options.dataBorder;

                    if (!border || typeof border.pt !== "number" || !(border.pt > 0) || !Number.isFinite(border.pt) || typeof border.color !== "string" || border.color === "") {
                        unsupported.push(`series[${index}].options.dataBorder[needs a positive pt and a color]`);
                    }
                }
            }
        }

        if (unsupported.length) {
            throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: unsupported.sort() });
        }
    }

    const requestedValueAxes = options.valAxes ?? [];
    const unsupportedAxisOptions: string[] = [];
    const requestedCategoryAxes = options.catAxes ?? [];

    // PptxGenJS 4.0.1 throws without as many valAxes as catAxes.
    if (requestedCategoryAxes.length > 0 && requestedCategoryAxes.length !== requestedValueAxes.length) {
        throw new TypeError(`catAxes must have as many entries as valAxes; got ${requestedCategoryAxes.length} and ${requestedValueAxes.length}`);
    }

    // A combo type on the secondary category axis: PptxGenJS 4.0.1 writes catAxes[1] as that axis (#603).
    const secondaryCategoryTypes = mixed && type.some(chart => chart.options?.secondaryCatAxis === true);

    for (const [index, axis] of requestedCategoryAxes.entries()) {
        for (const key of Object.keys(axis)) {
            // catAxes[1] takes the labels' and title's options where a type draws on it (SECONDARY_CATEGORY_AXIS_OPTIONS).
            // catAxisCrossesAt is read from the value axes and catLabelFormatCode from the chart (its date axis and its
            // categories' number format would part), so neither is taken from catAxes[0].
            const supported = index === 0 ? CATEGORY_AXIS_OPTION_SET.has(key) && key !== "catAxisCrossesAt" && key !== "catLabelFormatCode" : secondaryCategoryTypes && SECONDARY_CATEGORY_AXIS_OPTIONS.has(key);

            if (axis[key as keyof typeof axis] !== undefined && !supported) {
                unsupportedAxisOptions.push(`catAxes[${index}].${key}`);
            }
        }
    }

    for (const [index, axis] of requestedValueAxes.entries()) {
        for (const key of Object.keys(axis)) {
            if (axis[key as keyof PptxGenJS.IChartPropsAxisVal] !== undefined && !SUPPORTED_VALUE_AXIS_OPTIONS.has(key)) {
                unsupportedAxisOptions.push(`valAxes[${index}].${key}`);
            }
        }
    }

    for (const [index, axis] of requestedValueAxes.entries()) {
        for (const key of Object.keys(axis.valGridLine ?? {})) {
            if (!SUPPORTED_GRID_LINE_OPTIONS.has(key)) {
                unsupportedAxisOptions.push(`valAxes[${index}].valGridLine.${key}`);
            }
        }
    }

    for (const key of Object.keys(options.valGridLine ?? {})) {
        if (!SUPPORTED_GRID_LINE_OPTIONS.has(key)) {
            unsupportedAxisOptions.push(`valGridLine.${key}`);
        }
    }

    for (const key of Object.keys(options.catGridLine ?? {})) {
        if (!SUPPORTED_GRID_LINE_OPTIONS.has(key)) {
            unsupportedAxisOptions.push(`catGridLine.${key}`);
        }
    }

    for (const key of Object.keys(options.serGridLine ?? {})) {
        if (!SUPPORTED_GRID_LINE_OPTIONS.has(key)) {
            unsupportedAxisOptions.push(`serGridLine.${key}`);
        }
    }

    // A scatter or bubble chart's table is drawn as none where the target draws none (drawsScatterDataTable).
    if (options.showDataTable && !mixed && XY_TYPES.has(type as NormalizedChartType) && profile.drawsScatterDataTable) {
        unsupportedAxisOptions.push(`showDataTable[${String(type)}]`);
    }

    // PptxGenJS writes no axes on a pie or doughnut, so the category-axis options draw nothing there. A minMax
    // orientation and an autoZero crossing are the defaults, drawn as they are.
    const undrawnCategoryOptions = mixed ? undefined : CATEGORY_AXIS_UNDRAWN_OPTIONS[type as NormalizedChartType];

    for (const key of undrawnCategoryOptions ?? []) {
        const value: unknown = options[key as keyof PptxGenJS.IChartOpts];

        if (value !== undefined && value !== "minMax" && value !== "autoZero") {
            unsupportedAxisOptions.push(`${key}[${String(type)}]`);
        }
    }

    // bar3D draws its data labels without a box.
    if (type === "bar3D" && options.dataLabelBkgrdColors) {
        unsupportedAxisOptions.push("dataLabelBkgrdColors[bar3D]");
    }

    // The value-axis options a chart type does not draw, on the chart or a valAxes entry.
    const undrawnValueOptions = mixed ? undefined : VALUE_AXIS_UNDRAWN_OPTIONS[type as NormalizedChartType];
    const valueAxisOptionSets: [string, PptxGenJS.IChartPropsAxisVal][] = [["", options]];

    for (const [index, axis] of requestedValueAxes.entries()) {
        valueAxisOptionSets.push([`valAxes[${index}].`, axis]);
    }

    for (const [prefix, axis] of valueAxisOptionSets) {
        for (const key of undrawnValueOptions ?? []) {
            const value = axis[key as keyof PptxGenJS.IChartPropsAxisVal];

            // A minMax orientation and a 0 rotation are the defaults, drawn as they are.
            if (value !== undefined && value !== "minMax" && !(key === "valAxisLabelRotate" && value === 0)) {
                unsupportedAxisOptions.push(`${prefix}${key}[${String(type)}]`);
            }
        }

        // A radar chart draws its value labels beside the first spoke or none.
        if (type === "radar" && axis.valAxisLabelPos !== undefined && axis.valAxisLabelPos !== "nextTo" && axis.valAxisLabelPos !== "none") {
            unsupportedAxisOptions.push(`${prefix}valAxisLabelPos[${axis.valAxisLabelPos}][radar]`);
        }
    }

    if (type === "bar3D" && options.barDir === "bar") {
        unsupportedAxisOptions.push("barDir[bar3D]");
    }

    // The deprecated top-level border is checked as plotArea.border is.
    for (const [key, value] of Object.entries({ ...options.plotArea, ...(options.border ? { border: options.border } : {}) })) {
        const supported = SUPPORTED_PLOT_AREA_OPTIONS[key];

        if (!supported) {
            unsupportedAxisOptions.push(`plotArea.${key}`);
            continue;
        }

        for (const nested of Object.keys(value ?? {})) {
            if (!supported.has(nested)) {
                unsupportedAxisOptions.push(`plotArea.${key}.${nested}`);
            }
        }
    }

    if (!mixed && !profile.drawsPlotAreaOnEveryType && PLOT_AREA_UNDRAWN_TYPES.has(type as NormalizedChartType)) {
        for (const key of ["plotArea", "fill", "border"] as const) {
            if (options[key] !== undefined) {
                unsupportedAxisOptions.push(`${key}[${String(type)}]`);
            }
        }
    }

    if (unsupportedAxisOptions.length) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: unsupportedAxisOptions.sort() });
    }

    const unsupportedOptions: string[] = [];

    for (const key of Object.keys(options)) {
        if (options[key as keyof PptxGenJS.IChartOpts] !== undefined && !SUPPORTED_CHART_OPTION_SET.has(key)) {
            unsupportedOptions.push(key);
        }
    }

    unsupportedOptions.sort();

    if (unsupportedOptions.length) {
        throw new UnsupportedChartError({
            reason: "chart-option",
            chartTypes: [mixed ? "mixed" : type],
            unsupportedOptions,
        });
    }

    const valueAxisOptions = options.valAxes ?? [];
    const normalizedType: NormalizedChartType = mixed ? "mixed" : (type as NormalizedChartType);
    const palette = options.chartColors?.length ? options.chartColors : normalizedType === "pie" || normalizedType === "doughnut" ? PPTX_DEFAULTS.chart.colors.pie : PPTX_DEFAULTS.chart.colors.bar;
    // PptxGenJS 4.0.1 writes a `c:dPt` per point of a bar or bar3D group's only series when the caller's chartColors has
    // more than one entry: a bar's is filled with `chartColors[index]`, a bar3D's has only an outline (`<a:ln>`) in it,
    // and with `lineSize: 0` either has only `<a:ln><a:noFill/>`. PowerPoint and LibreOffice draw each filled bar in its
    // own color and key the legend entry with the series color (#473); a point without a fill takes the automatic one,
    // where the target draws it (chartPointKeepsSeriesFill, #507).
    // invertedColors writes the same points too, a negative one in `invertedColors[index]` (LibreOffice 26.2 draws each in
    // it, a positive one in the palette's entry, #567).
    const invertedColors = options.invertedColors?.length ? options.invertedColors : undefined;
    const pointColors = (options.chartColors && options.chartColors.length > 1) || invertedColors ? palette : undefined;
    // A combo chart type's own invertedColors replace the chart's on its series (#594); empty: none.
    const pointColor = (colors: readonly string[], inverted: readonly string[], value: number, index: number): CssColor => {
        const source = value < 0 && inverted.length > 0 ? inverted : colors;

        return normalizeColor(source[index % source.length] ?? "");
    };
    const withPointStyle = (item: NormalizedChartSeries, colors: readonly string[], inverted: readonly string[]): NormalizedChartSeries => {
        // A bar type reads the chart's lineSize, not a line type's.
        if (item.type === "bar" && requestedOptions.lineSize !== 0) {
            // A combo chart type's own palette is not the chart's, so each point names its fill.
            return { ...item, varyColors: true, ...(inverted.length > 0 || colors !== palette ? { pointFills: item.values.map((value, index) => pointColor(colors, inverted, value, index)) } : {}) };
        }

        // shortcut: the automatic fill of the first series; a single-series bar group after other series in a combo
        // chart may take a later accent, measure it when a deck needs it.
        const automaticPointFill = normalizeColor(PPTX_DEFAULTS.theme.colors.accent1);

        if (item.type === "bar3D" && requestedOptions.lineSize !== 0) {
            return { ...item, automaticPointFill, pointLineColors: item.values.map((value, index) => pointColor(colors, inverted, value, index)) };
        }

        return { ...item, automaticPointFill };
    };
    // dataBorder: `<a:ln w="${pt || 0.75}">` in its color on bar, bar3D and area series, bubbles and pie slices; a line,
    // radar or scatter series has its own line.
    const dataBorder: NormalizedChart["dataBorder"] =
        options.dataBorder && DATA_BORDER_TYPES.has(normalizedType)
            ? {
                  color: normalizeColor(options.dataBorder.color || PPTX_DEFAULTS.chart.dataBorder.color),
                  width: checkedSize(options.dataBorder.pt || PPTX_DEFAULTS.chart.dataBorder.widthPt, "dataBorder.pt", PPTX_DEFAULTS.limits.lineWidthPt),
              }
            : undefined;
    // A combo chart's lines are smoothed as its first line type says, where the target reads it so (#594).
    const firstLineType = mixed ? type.find(chart => chart.type === "line") : undefined;
    const lineSmooth = (profile.comboLineSmoothFromFirstLineType ? firstLineType?.options?.lineSmooth : undefined) ?? options.lineSmooth ?? PPTX_DEFAULTS.chart.lineSmooth;
    // The chart's own style: `lineDataSymbolLineSize || 0.75` in points (addChartDefinition converts it to EMU).
    const chartStyle = normalizeSeriesStyle(options, dataBorder, options.lineDataSymbolLineSize || PPTX_DEFAULTS.chart.lineDataSymbolLineSizePt, lineSmooth);
    let series: NormalizedChartSeries[];
    // A combo chart series' chart type options, by series index.
    const seriesOptions: PptxGenJS.IChartOpts[] = [];
    // A combo chart type's own chartColors, after the chart's palette: its series index them from their offset.
    const typePalette: string[] = [];

    if (!Array.isArray(type)) {
        series = normalizeSeries(normalizedType, data);

        if ((normalizedType === "bar" || normalizedType === "bar3D") && pointColors && series.length === 1 && series[0]) {
            series[0] = withPointStyle(series[0], pointColors, invertedColors ?? []);
        }
    } else {
        series = [];

        for (const chart of type) {
            // PptxGenJS restarts colors within each mixed chart-type group, from the type's own chartColors when it has
            // them (`{ ...options, ...type.options }`).
            const ownColors = chart.options?.chartColors?.length ? chart.options.chartColors : undefined;
            const colorOffset = palette.length + typePalette.length;

            if (ownColors) {
                typePalette.push(...ownColors);
            }

            const typeInverted = chart.options?.invertedColors?.length ? chart.options.invertedColors : (invertedColors ?? []);
            const typePointColors = ownColors ? (ownColors.length > 1 || typeInverted.length > 0 ? ownColors : undefined) : chart.options?.invertedColors?.length ? palette : pointColors;
            // The type's own fill, outline, line and marker options make its series' style (#594). PptxGenJS writes them
            // unchecked: a lineDataSymbolLineSize as EMU (addChartDefinition converts only the chart's), and the
            // dataBorder checked above. A target may smooth every line as the first line type says.
            const ownStyle = MIXED_STYLE_OPTIONS.some(key => chart.options?.[key] !== undefined);
            const typeBorder = chart.options?.dataBorder;
            const typeLineSymbolLineSize = chart.options?.lineDataSymbolLineSize;
            const style = ownStyle
                ? normalizeSeriesStyle(
                      { ...options, ...chart.options },
                      typeBorder ? { color: normalizeColor(typeBorder.color ?? ""), width: checkedSize(typeBorder.pt, "dataBorder.pt", PPTX_DEFAULTS.limits.lineWidthPt) } : chartStyle.dataBorder,
                      typeLineSymbolLineSize === undefined ? chartStyle.lineSymbolLineSize : typeLineSymbolLineSize / PPTX_DEFAULTS.units.emuPerPoint,
                      profile.comboLineSmoothFromFirstLineType ? lineSmooth : (chart.options?.lineSmooth ?? options.lineSmooth ?? PPTX_DEFAULTS.chart.lineSmooth),
                  )
                : undefined;

            for (const item of normalizeSeries(chart.type as NormalizedChartType, chart.data)) {
                const colorIndex = ownColors ? colorOffset + (item.colorIndex % ownColors.length) : item.colorIndex;
                // PptxGenJS fills a line type's markers with `chartColors[index of the series in the chart]` (a random entry
                // past the end) and outlines them in the line's color; a target may read that fill (readsComboMarkerFill, #604).
                // shortcut: past the end of chartColors the fill is random, so the markers keep the line's color.
                const markerColor = chart.type === "line" && profile.readsComboMarkerFill ? (ownColors ?? palette)[series.length] : undefined;
                const lineColor = [...palette, ...typePalette][colorIndex % (palette.length + typePalette.length)];
                const markerFill = markerColor !== undefined && lineColor !== undefined && normalizeColor(markerColor) !== normalizeColor(lineColor) ? normalizeColor(markerColor) : undefined;
                // A secondary category type's series run along that axis where the target draws them so (#603).
                const onSecondaryCategoryAxis = chart.options?.secondaryCatAxis === true && profile.secondaryCategoryAxis.seriesFollowAxis;
                const placed = {
                    ...item,
                    valueAxisIndex: chart.options?.secondaryValAxis ? (1 as const) : (0 as const),
                    ...(onSecondaryCategoryAxis ? { onSecondaryCategoryAxis } : {}),
                    colorIndex,
                    ...(style ? { style } : {}),
                    ...(markerFill ? { markerFill } : {}),
                };

                series.push(chart.type === "bar" && chart.data.length === 1 && typePointColors ? withPointStyle(placed, typePointColors, typeInverted) : placed);
                seriesOptions.push(chart.options ?? {});
            }
        }
    }

    // Refuse series the chart type cannot draw.
    if ((normalizedType === "pie" || normalizedType === "doughnut") && series.length !== 1) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["data[multiple-series]"] });
    }

    if ((normalizedType === "scatter" || normalizedType === "bubble") && series.length < 2) {
        throw new TypeError(`${normalizedType} charts require an X-value series followed by at least one Y-value series`);
    }

    if (normalizedType === "bubble" && series.slice(1).some(item => !item.sizes)) {
        throw new TypeError("Bubble Y-value series must include sizes");
    }

    if (!["scatter", "bubble", "pie", "doughnut"].includes(normalizedType)) {
        const labels = JSON.stringify(series[0]?.labels ?? []);

        if (series.some(item => JSON.stringify(item.labels) !== labels)) {
            throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["data.labels[per-series]"] });
        }
    }

    // The grouping PptxGenJS 4.0.1 writes per chart type: bar keeps clustered, stacked or percentStacked (standard and
    // anything else become clustered), bar3D any of the four (else standard), and an area chart `<c:grouping>` only for
    // stacked; a line or any other chart gets none, so it is drawn standard (addChartDefinition, makeChartType).
    const grouping = options.barGrouping;
    let normalizedGrouping: NormalizedChart["grouping"] = PPTX_DEFAULTS.chart.grouping.none;

    if (normalizedType === "bar") {
        normalizedGrouping = grouping === "stacked" || grouping === "percentStacked" || grouping === "clustered" ? grouping : PPTX_DEFAULTS.chart.grouping.bar;
    } else if (normalizedType === "bar3D") {
        normalizedGrouping = grouping === "stacked" || grouping === "percentStacked" || grouping === "standard" || grouping === "clustered" ? grouping : PPTX_DEFAULTS.chart.grouping.bar3D;
    } else if (normalizedType === "area") {
        normalizedGrouping = grouping === "stacked" ? grouping : PPTX_DEFAULTS.chart.grouping.area;
    } else if (normalizedType === "mixed") {
        // The bar group's `barGrouping || 'clustered'`: addChartDefinition checks no grouping on a combo chart.
        normalizedGrouping = grouping === "stacked" || grouping === "percentStacked" || grouping === "standard" || grouping === "clustered" ? grouping : PPTX_DEFAULTS.chart.grouping.bar;
    }

    // Each series stacks as its chart element's <c:grouping> says, also in a combo chart, where PptxGenJS writes one
    // element per type: the bar grouping on the bars, "stacked" on the areas only for "stacked", none on the lines (#529).
    // A target stacks a combo chart's percentStacked bars by value or as shares of 1 (comboPercentBarStack, #537).
    // A combo chart type's own barGrouping replaces the chart's on its element, unchecked (#590).
    const seriesGroupings: NormalizedChart["grouping"][] = [];

    for (const [index, item] of series.entries()) {
        const typeGrouping = seriesOptions[index]?.barGrouping ?? grouping;
        const itemGrouping = !mixed ? normalizedGrouping : typeGrouping === "stacked" || typeGrouping === "percentStacked" || typeGrouping === "standard" || typeGrouping === "clustered" ? typeGrouping : PPTX_DEFAULTS.chart.grouping.bar;
        const barStack = mixed && itemGrouping === "percentStacked" ? profile.comboPercentBarStack : itemGrouping === "stacked" || itemGrouping === "percentStacked" ? itemGrouping : undefined;
        const stack = item.type === "line" ? undefined : item.type === "area" ? (itemGrouping === "stacked" ? itemGrouping : undefined) : barStack;

        seriesGroupings.push(itemGrouping);

        if (stack) {
            series[index] = { ...item, stack };
        }
    }

    const stackedBars = series.some(item => item.type === "bar" && item.stack !== undefined);

    const hasSecondaryValueAxis = series.some(item => item.valueAxisIndex === 1);

    if (hasSecondaryValueAxis && valueAxisOptions.length !== 2) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["valAxes[primary,secondary]"] });
    }

    if (!hasSecondaryValueAxis && valueAxisOptions.length > 1) {
        throw new UnsupportedChartError({ reason: "chart-combination", chartTypes, unsupportedOptions: ["valAxes[secondary-without-series]"] });
    }

    // The data labels: a combo chart's chart types each take their own (`{ ...options, ...type.options }`).
    const dataLabels = normalizeDataLabels(normalizedType, options, normalizedGrouping, chartTypes, profile);

    if (mixed) {
        for (const [index, item] of series.entries()) {
            series[index] = { ...item, dataLabels: normalizeDataLabels(item.type, { ...options, ...seriesOptions[index] }, seriesGroupings[index] ?? normalizedGrouping, chartTypes, profile) };
        }
    }

    const axisSeries = (axisIndex: 0 | 1) => series.filter(item => item.valueAxisIndex === axisIndex);
    const primaryDataRange = valueAxisDataRange(normalizedType, axisSeries(0));
    const secondaryDataRange = hasSecondaryValueAxis ? valueAxisDataRange(normalizedType, axisSeries(1)) : undefined;
    const catDefaults = PPTX_DEFAULTS.chart.categoryAxis;
    // PptxGenJS 4.0.1 gives a scatter chart (not a bubble chart) `{ color: 'D9D9D9', size: 1 }` gridlines on both axes
    // when catGridLine or valGridLine is unset.
    const scatterGridLine: PptxGenJS.OptsChartGridLine = { color: PPTX_DEFAULTS.chart.scatterGridLine.color, size: PPTX_DEFAULTS.chart.scatterGridLine.widthPt };
    const catGridLine = options.catGridLine ?? (normalizedType === "scatter" ? scatterGridLine : undefined);

    // The category axis's tick marks, crossing, label skip and number format (makeCatAxis, #566).
    const catMajorTickMark = tickMark(options.catAxisMajorTickMark, "catAxisMajorTickMark", PPTX_DEFAULTS.chart.categoryMajorTickMark);
    const catMinorTickMark = tickMark(options.catAxisMinorTickMark, "catAxisMinorTickMark", catDefaults.minorTickMark);
    const categoryCount = series[0]?.labels.length ?? 0;
    // `<c:crossesAt val="${catAxisCrossesAt}"/>` on the value axis is the 1-based category it crosses before; a string
    // is written `<c:crosses val="...">` ("autoZero", "min" or "max").
    const catCrossesAt: unknown = options.catAxisCrossesAt;
    let valueAxisCrossing: NormalizedChart["categoryAxis"]["valueAxisCrossing"];

    if (typeof catCrossesAt === "number" && !Number.isFinite(catCrossesAt)) {
        throw new TypeError(`catAxisCrossesAt must be a finite category number or "autoZero"; got ${catCrossesAt}`);
    }

    if (typeof catCrossesAt === "number" && XY_TYPES.has(normalizedType)) {
        // A scatter or bubble chart's x axis is a value axis: the y axis crosses it at that x value (#591).
        valueAxisCrossing = { kind: "value", value: catCrossesAt };
    } else if (typeof catCrossesAt === "number") {
        // Before the first category is the default; past the last is after it.
        const index = Math.min(Math.max(0, Math.round(catCrossesAt) - 1), categoryCount);

        valueAxisCrossing = index === 0 ? undefined : index >= categoryCount ? { kind: "max" } : { kind: "category", index };
    } else if (catCrossesAt === "max") {
        valueAxisCrossing = { kind: "max" };
    } else if (catCrossesAt !== undefined && catCrossesAt !== "autoZero" && catCrossesAt !== "min") {
        throw new TypeError(`catAxisCrossesAt must be a category number, "autoZero", "min" or "max"; got ${JSON.stringify(catCrossesAt)}`);
    }

    const valCrossesAt: unknown = options.valAxisCrossesAt;
    let categoryAxisCrossing: NormalizedChartValueAxis["categoryAxisCrossing"];

    if (typeof valCrossesAt === "number") {
        if (!Number.isFinite(valCrossesAt)) {
            throw new TypeError(`valAxisCrossesAt must be a finite value or "autoZero"; got ${valCrossesAt}`);
        }

        categoryAxisCrossing = { kind: "value", value: valCrossesAt };
    } else if (valCrossesAt === "min" || valCrossesAt === "max") {
        categoryAxisCrossing = { kind: valCrossesAt };
    } else if (valCrossesAt !== undefined && valCrossesAt !== "autoZero") {
        throw new TypeError(`valAxisCrossesAt must be a value, "autoZero", "min" or "max"; got ${JSON.stringify(valCrossesAt)}`);
    }

    // `<c:tickLblSkip val="${catAxisLabelFrequency}"/>`, read by the target or not (readsCategoryLabelSkip).
    let labelSkip: number | undefined;

    if (options.catAxisLabelFrequency !== undefined) {
        const skip = Number(options.catAxisLabelFrequency);

        if (!Number.isInteger(skip) || skip < 1) {
            throw new TypeError(`catAxisLabelFrequency must be a whole number of categories from 1; got ${JSON.stringify(options.catAxisLabelFrequency)}`);
        }

        // LibreOffice 26.2.6 skips no scatter or bubble x label (a value axis) and no radar spoke name.
        labelSkip = profile.readsCategoryLabelSkip && skip > 1 && !XY_TYPES.has(normalizedType) ? skip : undefined;
    }

    // A scatter or bubble chart's x axis keeps its General numbers in LibreOffice 26.2.6 whatever catLabelFormatCode says.
    const requestedFormatCode = XY_TYPES.has(normalizedType) ? undefined : options.catLabelFormatCode;
    const timeUnits: Partial<Record<"catAxisBaseTimeUnit" | "catAxisMajorTimeUnit" | "catAxisMinorTimeUnit", ChartTimeUnit>> = {};
    const axisNumbers: Partial<Record<"catAxisMajorUnit" | "catAxisMinorUnit" | "catAxisMinVal" | "catAxisMaxVal", number>> = {};

    for (const key of ["catAxisBaseTimeUnit", "catAxisMajorTimeUnit", "catAxisMinorTimeUnit"] as const) {
        const unit: unknown = options[key];

        if (unit !== undefined && !TIME_UNITS.has(String(unit))) {
            throw new TypeError(`${key} must be "days", "months" or "years"; got ${JSON.stringify(unit)}`);
        }

        if (unit !== undefined) {
            // TIME_UNITS holds exactly the ChartTimeUnit names.
            timeUnits[key] = unit as ChartTimeUnit;
        }
    }

    for (const key of ["catAxisMajorUnit", "catAxisMinorUnit", "catAxisMinVal", "catAxisMaxVal"] as const) {
        const value: unknown = options[key];

        if (value !== undefined && (typeof value !== "number" || !Number.isFinite(value) || (key.endsWith("Unit") && value <= 0))) {
            throw new TypeError(`${key} must be a finite${key.endsWith("Unit") ? " positive" : ""} number; got ${JSON.stringify(value)}`);
        }

        if (typeof value === "number") {
            axisNumbers[key] = value;
        }
    }

    // PptxGenJS 4.0.1 writes a `<c:dateAx>` for catLabelFormatCode, its categories as numbers. LibreOffice 26.2.6 lays
    // date serials out on a time line (dateAxisLayout); the time units and catAxisMajorUnit and catAxisMinorUnit are
    // written with catLabelFormatCode only. catAxisMinVal and catAxisMaxVal on a text axis, and catAxisMultiLevelLabels on
    // single-level labels (`<c:noMultiLvlLbl>`), change nothing there. shortcut: catAxisMinorUnit and
    // catAxisMinorTimeUnit draw no minor ticks between the slots yet; draw them when a fixture shows them.
    const serials = (series[0]?.labels ?? []).map(label => (label.trim() === "" ? NaN : Number(label)));
    const dateAxis = requestedFormatCode !== undefined && DATE_AXIS_TYPES.has(normalizedType) && serials.length > 0 && serials.every(serial => Number.isFinite(serial) && serial >= 1);
    const dateFormat = requestedFormatCode !== undefined && isSupportedChartDateFormat(requestedFormatCode);

    if (requestedFormatCode !== undefined && !isSupportedChartNumberFormat(requestedFormatCode) && !(dateAxis && dateFormat)) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["catLabelFormatCode[format]"] });
    }

    if (dateAxis && axisNumbers.catAxisMajorUnit !== undefined && !Number.isInteger(axisNumbers.catAxisMajorUnit)) {
        throw new TypeError(`catAxisMajorUnit must be a whole number of time units on a date axis; got ${axisNumbers.catAxisMajorUnit}`);
    }

    let dateSlots: NormalizedChart["categoryAxis"]["dateSlots"];

    // Consecutive whole serials with nothing else set are a slot each, the categories as they are.
    const consecutive = serials.every((serial, index) => Number.isInteger(serial) && (index === 0 || serial === (serials[index - 1] ?? 0) + 1));
    const timeLine = dateFormat || !consecutive || Object.keys(timeUnits).length > 0 || axisNumbers.catAxisMajorUnit !== undefined || axisNumbers.catAxisMinVal !== undefined || axisNumbers.catAxisMaxVal !== undefined;

    if (dateAxis && requestedFormatCode !== undefined && timeLine) {
        const layout = dateAxisLayout({
            serials,
            ...(timeUnits.catAxisBaseTimeUnit ? { baseUnit: timeUnits.catAxisBaseTimeUnit } : {}),
            ...(axisNumbers.catAxisMajorUnit !== undefined ? { majorUnit: axisNumbers.catAxisMajorUnit } : {}),
            ...(timeUnits.catAxisMajorTimeUnit ? { majorTimeUnit: timeUnits.catAxisMajorTimeUnit } : {}),
            ...(axisNumbers.catAxisMinVal !== undefined ? { minimum: axisNumbers.catAxisMinVal } : {}),
            ...(axisNumbers.catAxisMaxVal !== undefined ? { maximum: axisNumbers.catAxisMaxVal } : {}),
        });

        // LibreOffice 26.2.6 draws the points of one slot (daily dates on a months base unit) at one x (#593).
        if (layout.kind === "shared-slot") {
            throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: [`catAxisBaseTimeUnit[several dates in one ${layout.unit.slice(0, -1)}]`] });
        }

        if (layout.kind === "too-many-slots") {
            throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: [`catLabelFormatCode[time line of ${layout.count} slots, at most ${DATE_AXIS_SLOT_LIMIT}]`] });
        }

        if (layout.labelSerials.length === 0) {
            throw new TypeError("catAxisMinVal must not be after catAxisMaxVal on a date axis");
        }

        // A data table on a time line of a labelled slot per category, in order, names each column by its slot's date:
        // LibreOffice 26.2.6 draws Jan-24..Apr-24 over monthly serials with `mmm-yy` (#593). shortcut: one with empty
        // slots or a slot shared by several dates is not measured; measure it when a deck needs one.
        const slotPerCategory = layout.labelSerials.length === serials.length && layout.labelSerials.every(serial => serial !== null) && layout.categorySlots.every((slot, index) => slot === index);

        if (options.showDataTable && !slotPerCategory) {
            throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["showDataTable[date axis with empty slots]"] });
        }

        dateSlots = {
            labels: layout.labelSerials.map(serial => (serial === null ? "" : dateFormat ? formatChartDate(serial, requestedFormatCode) : formatChartNumber(serial, requestedFormatCode))),
            categorySlots: layout.categorySlots,
        };
    }

    // A time line's labels are formatted already.
    const labelFormatCode = dateSlots ? undefined : requestedFormatCode;
    // A scatter or bubble chart's x axis takes catAxisMinVal, catAxisMaxVal and catAxisMajorUnit as a value axis's bounds
    // and step (PptxGenJS writes them on its c:valAx); shortcut: catAxisMinorUnit draws no minor ticks there yet.
    const xScale = XY_TYPES.has(normalizedType)
        ? {
              ...(axisNumbers.catAxisMinVal !== undefined ? { minimum: axisNumbers.catAxisMinVal } : {}),
              ...(axisNumbers.catAxisMaxVal !== undefined ? { maximum: axisNumbers.catAxisMaxVal } : {}),
              ...(axisNumbers.catAxisMajorUnit !== undefined ? { majorUnit: axisNumbers.catAxisMajorUnit } : {}),
          }
        : {};

    // A category axis's title and labels, from the chart's options merged with its catAxes entry (#594, #603).
    const categoryAxisText = (source: PptxGenJS.IChartOpts) => {
        const sourceOrientation: string | undefined = source.catAxisOrientation;

        if (sourceOrientation !== undefined && sourceOrientation !== "minMax" && sourceOrientation !== "maxMin") {
            throw new TypeError(`catAxisOrientation must be "minMax" or "maxMin"; got ${JSON.stringify(sourceOrientation)}`);
        }

        return {
            reversed: sourceOrientation === "maxMin",
            showTitle: source.showCatAxisTitle ?? false,
            title: source.catAxisTitle || catDefaults.title,
            titleColor: normalizeColor(source.catAxisTitleColor || catDefaults.titleColor),
            titleFontFace: source.catAxisTitleFontFace || PPTX_DEFAULTS.chart.fontFace,
            titleFontSize: positiveNumber(source.catAxisTitleFontSize, "catAxisTitleFontSize", catDefaults.titleFontSizePt),
            titleRotate: source.catAxisTitleRotate === undefined ? undefined : finiteNumber(source.catAxisTitleRotate, "catAxisTitleRotate", 0),
            labelColor: normalizeColor(source.catAxisLabelColor || catDefaults.labelColor),
            labelFontFace: source.catAxisLabelFontFace || PPTX_DEFAULTS.chart.fontFace,
            labelFontSize: positiveNumber(source.catAxisLabelFontSize, "catAxisLabelFontSize", catDefaults.labelFontSizePt),
            labelBold: source.catAxisLabelFontBold ?? catDefaults.labelBold,
            labelItalic: source.catAxisLabelFontItalic ?? catDefaults.labelItalic,
            hidden: source.catAxisHidden ?? catDefaults.hidden,
        };
    };
    // PptxGenJS 4.0.1 writes catAxes[1] from `Object.assign({}, options, catAxes[1])`, so it keeps every chart-level
    // catAxis* option it does not set (#603). The target draws its labels in a row of their own or over the primary ones.
    // shortcut: the secondary axis line takes no option (catAxisLineColor and the others style the primary line only);
    // draw them when a deck styles the secondary line.
    const secondaryCategoryOptions: PptxGenJS.IChartOpts = { ...requestedOptions, ...requestedOptions.catAxes?.[1] };
    const { reversed: secondaryReversed, ...secondaryText } = categoryAxisText(secondaryCategoryOptions);
    const secondaryCategoryAxis: NormalizedChart["secondaryCategoryAxis"] = secondaryCategoryTypes
        ? {
              ...secondaryText,
              ...(secondaryReversed ? { reversed: true } : {}),
              labelPosition: secondaryCategoryOptions.catAxisLabelPos ?? catDefaults.columnLabelPosition,
              placement: profile.secondaryCategoryAxis.placement,
          }
        : undefined;
    const { reversed, ...primaryText } = categoryAxisText(options);
    const categoryAxis: NormalizedChart["categoryAxis"] = {
        ...primaryText,
        // PptxGenJS writes `nextTo` on a scatter chart's x axis whatever catAxisLabelPos says.
        labelPosition: normalizedType === "scatter" ? "nextTo" : (options.catAxisLabelPos ?? (options.barDir === "bar" ? catDefaults.barLabelPosition : catDefaults.columnLabelPosition)),
        labelRotate: finiteNumber(options.catAxisLabelRotate, "catAxisLabelRotate", catDefaults.labelRotate),
        ...(labelSkip ? { labelSkip } : {}),
        ...(labelFormatCode !== undefined ? { labelFormatCode } : {}),
        ...(dateSlots ? { dateSlots } : {}),
        ...(Object.keys(xScale).length > 0 ? { xScale } : {}),
        ...(reversed ? { reversed: true } : {}),
        lineColor: normalizeColor(options.catAxisLineColor || catDefaults.lineColor),
        lineWidth: positiveNumber(options.catAxisLineSize || undefined, "catAxisLineSize", catDefaults.lineWidthPt),
        lineStyle: options.catAxisLineStyle ?? catDefaults.lineStyle,
        // PptxGenJS writes `<c:minorTickMark val="none"/>` on a scatter chart's axes, as its major ones.
        minorTickMark: normalizedType === "scatter" ? "none" : catMinorTickMark,
        ...(valueAxisCrossing ? { valueAxisCrossing } : {}),
        gridLineColor: normalizeColor(catGridLine?.color ?? PPTX_DEFAULTS.chart.gridLine.color),
        gridLineWidth: positiveNumber(catGridLine?.size, "catGridLine.size", PPTX_DEFAULTS.chart.gridLine.widthPt),
        gridLineStyle: catGridLine ? (catGridLine.style ?? PPTX_DEFAULTS.chart.gridLine.style) : catDefaults.gridLineStyle,
        gridLineCap: catGridLine?.cap ?? PPTX_DEFAULTS.chart.gridLine.cap,
    };
    const primaryScale: ValueAxisScaleKind = normalizedType === "bar3D" ? { kind: "bar3D" } : normalizedType === "radar" ? { kind: "radar", rounding: profile.radarScale.kind, maxIncrements: profile.radarScale.maxIncrements } : { kind: "linear" };
    // PptxGenJS 4.0.1 writes each value axis from `Object.assign({}, options, valAxes[i])`: a valAxes entry keeps every
    // chart-level valAxis* and valGridLine option it does not set itself (#596).
    const primaryAxis = normalizeValueAxis(
        valueAxisOptions[0] ? { ...options, ...valueAxisOptions[0] } : options,
        primaryDataRange,
        primaryScale,
        valueAxisOptions[0] ? "valAxes[0]." : "",
        chartTypes,
        normalizedType === "scatter" ? scatterGridLine : {},
        profile.readsValueAxisDisplayUnits,
    );
    // A bar3D value axis line and its tick marks are drawn where the target draws them (drawsBar3DValueAxisLine).
    const bar3DAxisUndrawn = normalizedType === "bar3D" && !profile.drawsBar3DValueAxisLine ? { lineVisible: false, majorTickMark: "none" as const, minorTickMark: "none" as const } : {};
    const valueAxes = [{ ...primaryAxis, ...bar3DAxisUndrawn, ...(categoryAxisCrossing ? { categoryAxisCrossing } : {}) }];

    // bar3D and radar charts have no axis band for the display unit's label (set only where the target reads units).
    if ((normalizedType === "bar3D" || normalizedType === "radar") && primaryAxis.displayUnit?.label) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: [`valAxisDisplayUnitLabel[${normalizedType}]`] });
    }

    // titlePos: PptxGenJS writes the title's `<c:manualLayout>` edge from the position plus the chart's own x or y in
    // inches, `t * (t / 5) / 10`, divided by 10 once when at least 1 and again when at least 0.1 (genXmlTitle).
    let titlePosition: NormalizedChart["titlePosition"];

    if (options.titlePos) {
        const chartX = options.x ?? PPTX_DEFAULTS.chart.x;
        const chartY = options.y ?? PPTX_DEFAULTS.chart.y;

        if (typeof chartX !== "number" || typeof chartY !== "number" || !Number.isFinite(options.titlePos.x) || !Number.isFinite(options.titlePos.y)) {
            throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["titlePos[percent chart position or non-numeric x, y]"] });
        }

        const edge = (total: number): number => {
            let value = total === 0 ? 0 : (total * (total / 5)) / 10;

            if (value >= 1) {
                value /= 10;
            }

            if (value >= 0.1) {
                value /= 10;
            }

            return value;
        };

        titlePosition = { x: edge(options.titlePos.x + chartX), y: edge(options.titlePos.y + chartY) };
    }

    const serDefaults = PPTX_DEFAULTS.chart.seriesAxis;
    const serOrientation: string | undefined = options.serAxisOrientation;

    if (serOrientation !== undefined && serOrientation !== "minMax" && serOrientation !== "maxMin") {
        throw new TypeError(`serAxisOrientation must be "minMax" or "maxMin"; got ${JSON.stringify(serOrientation)}`);
    }

    // `<c:tickLblSkip val="${serAxisLabelFrequency}"/>`, read by the target or not (readsSeriesLabelSkip).
    let seriesLabelSkip: number | undefined;

    if (options.serAxisLabelFrequency !== undefined) {
        const skip = Number(options.serAxisLabelFrequency);

        if (!Number.isInteger(skip) || skip < 1) {
            throw new TypeError(`serAxisLabelFrequency must be a whole number of series from 1; got ${JSON.stringify(options.serAxisLabelFrequency)}`);
        }

        seriesLabelSkip = profile.readsSeriesLabelSkip && skip > 1 ? skip : undefined;
    }

    // PptxGenJS defaults serGridLine to `{ style: "none" }` and replaces it whole, as valGridLine.
    const serGridLine = options.serGridLine ?? { style: "none" };
    // PptxGenJS writes the depth axis (makeSerAxis) on a bar3D chart only.
    const seriesAxis: NormalizedChart["seriesAxis"] =
        normalizedType === "bar3D"
            ? {
                  hidden: options.serAxisHidden ?? false,
                  reversed: serOrientation === "maxMin",
                  ...(seriesLabelSkip ? { labelSkip: seriesLabelSkip } : {}),
                  gridLineColor: normalizeColor(serGridLine.color ?? PPTX_DEFAULTS.chart.gridLine.color),
                  gridLineWidth: positiveNumber(serGridLine.size, "serGridLine.size", PPTX_DEFAULTS.chart.gridLine.widthPt),
                  gridLineStyle: serGridLine.style ?? PPTX_DEFAULTS.chart.gridLine.style,
                  gridLineCap: serGridLine.cap ?? PPTX_DEFAULTS.chart.gridLine.cap,
                  labelColor: normalizeColor(options.serAxisLabelColor || serDefaults.labelColor),
                  labelFontFace: options.serAxisLabelFontFace || PPTX_DEFAULTS.chart.fontFace,
                  labelFontSize: positiveNumber(options.serAxisLabelFontSize, "serAxisLabelFontSize", serDefaults.labelFontSizePt),
                  labelBold: options.serAxisLabelFontBold ?? false,
                  labelItalic: options.serAxisLabelFontItalic ?? false,
                  showTitle: options.showSerAxisTitle ?? false,
                  title: options.serAxisTitle || serDefaults.title,
                  titleColor: normalizeColor(options.serAxisTitleColor || serDefaults.titleColor),
                  titleFontFace: options.serAxisTitleFontFace || PPTX_DEFAULTS.chart.fontFace,
                  titleFontSize: positiveNumber(options.serAxisTitleFontSize, "serAxisTitleFontSize", serDefaults.titleFontSizePt),
                  ...(options.serAxisTitleRotate !== undefined ? { titleRotate: finiteNumber(options.serAxisTitleRotate, "serAxisTitleRotate", 0) } : {}),
              }
            : undefined;
    const tableDefaults = PPTX_DEFAULTS.chart.dataTable;
    // PptxGenJS 4.0.1 writes each `<c:showHorzBorder>`, `<c:showVertBorder>`, `<c:showOutline>` and `<c:showKeys>` 1
    // only when its option is true (its default-setting ternaries keep an unset option unset).
    // dataTableFormatCode is the values' `<c:numCache>` format, after valLabelFormatCode (which PptxGenJS 4.0.1 never
    // sets), read by the target or not (readsDataTableFormatCode).
    const tableFormatCode = profile.readsDataTableFormatCode ? options.dataTableFormatCode : undefined;

    if (tableFormatCode !== undefined && !isSupportedChartNumberFormat(tableFormatCode)) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["dataTableFormatCode[format]"] });
    }

    // LibreOffice 26.2.6 draws a bar3D chart's table borders and keys too, where it draws a 2D chart's text alone (#592).
    const tableBorders = profile.drawsDataTableBorders || normalizedType === "bar3D";
    const dataTable: NormalizedChart["dataTable"] =
        options.showDataTable && !DATA_TABLE_UNDRAWN_TYPES.has(normalizedType) && !XY_TYPES.has(normalizedType)
            ? {
                  ...(tableFormatCode !== undefined ? { valueFormatCode: tableFormatCode } : {}),
                  horizontalBorder: tableBorders && options.showDataTableHorzBorder === true,
                  verticalBorder: tableBorders && options.showDataTableVertBorder === true,
                  outline: tableBorders && options.showDataTableOutline === true,
                  keys: tableBorders && options.showDataTableKeys === true,
                  // shortcut: the theme's minor font (+mn-lt) is the default theme's; a presentation theme's body face is
                  // not passed to charts yet, pass it when a deck sets one with a data table.
                  fontFace: PPTX_DEFAULTS.theme.bodyFontFace,
                  fontSize: checkedSize(options.dataTableFontSize || tableDefaults.fontSizePt, "dataTableFontSize", PPTX_DEFAULTS.limits.fontSizePt),
                  color: normalizeColor(tableDefaults.color),
                  borderColor: normalizeColor(tableDefaults.borderColor),
                  borderWidth: tableDefaults.borderWidthPt,
              }
            : undefined;
    // `<c:overlap val="${stacked ? 100 : barOverlapPct || 0}"/>` on a bar chart; a stack is one bar already.
    const barOverlapPercent = (normalizedType === "bar" || normalizedType === "mixed") && !stackedBars ? Math.max(-100, Math.min(100, finiteNumber(options.barOverlapPct, "barOverlapPct", 0))) : 0;

    // Both targets take LibreOffice's measurements here (#595, #596, #591): PowerPoint cannot open PptxGenJS 4.0.1's XML
    // for a secondary value axis without a secondary category axis (dangling crossAx), so there is no PowerPoint reference.
    if (secondaryDataRange) {
        valueAxes.push(normalizeValueAxis({ ...options, ...valueAxisOptions[1] }, secondaryDataRange, { kind: "linear" }, "valAxes[1].", chartTypes, {}, profile.readsValueAxisDisplayUnits));
    }

    // PptxGenJS 4.0.1: a fill without a color writes <a:noFill/>; a border defaults a missing or zero pt and a missing
    // color to DEF_CHART_BORDER (1pt 363636). The deprecated top-level `fill` sets the fill's color and `border` replaces
    // the border.
    const plotFill = options.fill ? { ...options.plotArea?.fill, color: options.fill } : options.plotArea?.fill;
    const plotBorder = options.border ?? options.plotArea?.border;
    const plotArea: NormalizedChart["plotArea"] = {
        fill: normalizeFill(plotFill?.color ? plotFill : undefined),
        borderVisible: plotBorder !== undefined,
        borderColor: normalizeColor(plotBorder?.color || PPTX_DEFAULTS.chart.border.color),
        borderWidth: checkedSize(plotBorder?.pt || PPTX_DEFAULTS.chart.border.widthPt, "plotArea.border.pt", PPTX_DEFAULTS.limits.lineWidthPt),
    };
    // Each value that is not a number in 0..1 is dropped (with PptxGenJS's warning), then `|| default`: a 0 width or
    // height becomes the full frame, as PptxGenJS writes it.
    let plotLayout: NormalizedChart["plotLayout"];

    if (options.layout) {
        const resolved: NonNullable<NormalizedChart["plotLayout"]> = { ...PPTX_DEFAULTS.chart.plotLayout };

        for (const key of LAYOUT_KEYS) {
            const value = Number(options.layout[key]);

            if (Number.isNaN(value) || value < 0 || value > 1) {
                console.warn(`addChart: layout.${key} is ${JSON.stringify(options.layout[key])}, not a fraction of the chart frame from 0 to 1, so it is dropped and ${PPTX_DEFAULTS.chart.plotLayout[key]} is used, as PptxGenJS does.`);
                continue;
            }

            resolved[key] = value || PPTX_DEFAULTS.chart.plotLayout[key];
        }

        plotLayout = resolved;
    }

    return {
        type: normalizedType,
        series,
        colors: [...palette, ...typePalette].map(color => normalizeColor(color)),
        colorOpacity: chartStyle.colorOpacity,
        // Text options.
        showLegend: options.showLegend ?? PPTX_DEFAULTS.chart.showLegend,
        legendPosition: options.legendPos ?? PPTX_DEFAULTS.chart.legendPosition,
        // The rendering target places it (PuppeteerSlide, `topRightLegend`).
        topRightLegend: { kind: "beside" },
        showTitle: options.showTitle ?? PPTX_DEFAULTS.chart.showTitle,
        title: options.title ?? "Chart Title",
        // PptxGenJS 4.0.1 reads no chart-level `fontFace` or `fontSize`: every chart text has its own option,
        // `|| 'Arial'` and `|| 12` (the category-axis labels: `catAxisLabelFontSize || 12`).
        fontFace: PPTX_DEFAULTS.chart.fontFace,
        fontSize: PPTX_DEFAULTS.chart.fontSizePt,
        titleFontFace: options.titleFontFace ?? PPTX_DEFAULTS.chart.fontFace,
        titleFontSize: checkedSize(options.titleFontSize ?? PPTX_DEFAULTS.chart.titleFontSizePt, "titleFontSize", PPTX_DEFAULTS.limits.fontSizePt),
        titleBold: options.titleBold ?? false,
        titleColor: normalizeColor(options.titleColor ?? PPTX_DEFAULTS.text.color),
        ...(titlePosition ? { titlePosition } : {}),
        ...(options.titleRotate ? { titleRotate: finiteNumber(options.titleRotate, "titleRotate", 0) } : {}),
        // PptxGenJS never applies fontFace/fontSize to the legend.
        legendFontFace: options.legendFontFace ?? PPTX_DEFAULTS.theme.bodyFontFace,
        legendFontSize: checkedSize(options.legendFontSize ?? PPTX_DEFAULTS.chart.legendFontSizePt, "legendFontSize", PPTX_DEFAULTS.limits.fontSizePt),
        legendColor: normalizeColor(options.legendColor ?? PPTX_DEFAULTS.text.color),
        // Bar options.
        barDirection: options.barDir === "bar" ? ("bar" as const) : PPTX_DEFAULTS.chart.barDirection,
        // PptxGenJS 4.0.1 sets a stacked grouping's `barGapWidthPct` to 50 when it is unset or 0 (`if (!barGapWidthPct)`).
        barGapWidthPercent: Math.max(
            0,
            Math.min(1000, ((normalizedGrouping === "stacked" || normalizedGrouping === "percentStacked") && !options.barGapWidthPct ? PPTX_DEFAULTS.chart.stackedGapWidthPct : options.barGapWidthPct) ?? PPTX_DEFAULTS.chart.barGapWidthPct),
        ),
        ...(barOverlapPercent ? { barOverlapPercent } : {}),
        barGapDepthPercent: Math.max(0, Math.min(1000, options.barGapDepthPct ?? PPTX_DEFAULTS.chart.barGapDepthPct)),
        bar3DShape: BAR_3D_SHAPES.has(options.bar3DShape ?? "") ? (options.bar3DShape as NormalizedChart["bar3DShape"]) : PPTX_DEFAULTS.chart.bar3DShape,
        perspective3D: Math.max(0, Math.min(240, options.v3DPerspective ?? PPTX_DEFAULTS.chart.v3DPerspective)),
        rightAngleAxes3D: options.v3DRAngAx ?? PPTX_DEFAULTS.chart.v3DRightAngleAxes,
        rotationX3D: Math.max(-90, Math.min(90, options.v3DRotX ?? PPTX_DEFAULTS.chart.v3DRotationX)),
        rotationY3D: (((options.v3DRotY ?? PPTX_DEFAULTS.chart.v3DRotationY) % 360) + 360) % 360,
        // Doughnut, pie, radar and line options.
        doughnutHoleSize: Math.max(0, Math.min(90, options.holeSize ?? PPTX_DEFAULTS.chart.doughnutHoleSizePct)),
        firstSliceAngle: (((options.firstSliceAng ?? PPTX_DEFAULTS.chart.firstSliceAngle) % 360) + 360) % 360,
        radarStyle: options.radarStyle ?? PPTX_DEFAULTS.chart.radarStyle,
        lineSmooth: chartStyle.lineSmooth,
        lineSymbol: chartStyle.lineSymbol,
        lineSymbolSize: chartStyle.lineSymbolSize,
        ...(chartStyle.lineSymbolLineColor ? { lineSymbolLineColor: chartStyle.lineSymbolLineColor } : {}),
        lineSymbolLineSize: chartStyle.lineSymbolLineSize,
        lineSize: chartStyle.lineSize,
        dataLabels,
        plotArea,
        ...(plotLayout ? { plotLayout } : {}),
        // PptxGenJS writes `<c:crossBetween val="midCat"/>` on a combo chart with an area type (on an area chart alone,
        // "between"), so the categories sit on the ticks; the rendering target may keep a bar combo's between them.
        ...(mixed && type.some(chart => chart.type === "area") ? { categoriesOnTicks: true } : {}),
        altText: options.altText ?? "",
        grouping: normalizedGrouping,
        gridLineColor: normalizeColor(PPTX_DEFAULTS.chart.gridLine.color),
        gridLineWidth: PPTX_DEFAULTS.chart.gridLine.widthPt,
        axisLineVisible: options.catAxisLineShow ?? PPTX_DEFAULTS.chart.categoryAxisLineVisible,
        // PptxGenJS 4.0.1 writes `<c:majorTickMark val="none"/>` and `<c:minorTickMark val="none"/>` on both scatter
        // axes, whatever catAxisMajorTickMark or valAxisMajorTickMark say.
        categoryMajorTickMark: normalizedType === "scatter" ? "none" : catMajorTickMark,
        valueAxisMinimum: primaryAxis.minimum,
        valueAxisMaximum: primaryAxis.maximum,
        valueAxes: normalizedType === "scatter" ? valueAxes.map(axis => ({ ...axis, majorTickMark: "none" as const, minorTickMark: "none" as const })) : valueAxes,
        categoryAxis,
        ...(secondaryCategoryAxis ? { secondaryCategoryAxis } : {}),
        ...(seriesAxis ? { seriesAxis } : {}),
        ...(dataTable ? { dataTable } : {}),
        ...(dataBorder ? { dataBorder } : {}),
    };
}
