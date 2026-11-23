import type PptxGenJS from "pptxgenjs";
import { automaticAxisScale, bar3DAxisScale, type FixedAxisBound } from "#src/chart/axisScale";
import { UnsupportedChartError } from "#src/chart/errors";
import { isSupportedChartNumberFormat } from "#src/chart/format";
import { PPTX_DEFAULTS } from "#src/defaults";
import { DEFAULT_RENDERING_TARGET, RENDERING_TARGETS, type RenderingTarget, type RenderingTargetProfile } from "#src/target/index";
import type { CssColor, NormalizedChart, NormalizedChartDataLabels, NormalizedChartSeries, NormalizedChartType, NormalizedChartValueAxis } from "@pdflapas/model";
import { checkedSize } from "./size";
import { normalizeColor, normalizeFill } from "./style";

const SUPPORTED_TYPES = new Set<NormalizedChartType>(["area", "bar", "bar3D", "bubble", "doughnut", "line", "mixed", "pie", "radar", "scatter"]);

const MIXED_CHART_TYPES = new Set<PptxGenJS.CHART_NAME>(["area", "bar", "line"]);

export const CATEGORY_AXIS_OPTIONS = [
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
    "catAxisMinorTickMark",
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

/** The bar3D depth-axis options drawn: the series names on the depth axis and its title (#567). */
const SERIES_AXIS_OPTIONS = [
    "serAxisHidden",
    "serAxisLabelColor",
    "serAxisLabelFontBold",
    "serAxisLabelFontFace",
    "serAxisLabelFontItalic",
    "serAxisLabelFontSize",
    "serAxisOrientation",
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
    "v3DPerspective",
    "v3DRAngAx",
    "v3DRotX",
    "v3DRotY",
    "showValAxisTitle",
    "valAxes",
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

// LibreOffice 26.2.6 draws no plot-area fill or border on a pie, radar or bar3D chart (its PDF export of synthetic
// PptxGenJS 4.0.1 charts with `plotArea: { fill: { color: "FF0000" }, border: { color: "0000FF", pt: 3 } }`, #479),
// and PowerPoint is not measured, so they stay rejected there; `layout` places each of them.
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
const SUPPORTED_MIXED_SERIES_OPTIONS = new Set<string>(["secondaryValAxis", ...MIXED_LABEL_OPTIONS]);
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
// Where the axes cross, a reversed category axis, a label skip and a number format: drawn on a 2D category axis only.
const CATEGORY_AXIS_PLACEMENT_OPTIONS = ["catAxisCrossesAt", "catAxisLabelFrequency", "catAxisOrientation", "catLabelFormatCode", "valAxisCrossesAt"] as const;
// The category-axis options a chart type does not draw (#566). A radar chart's category axis is its spokes:
// LibreOffice 26.2.6 draws its labels in catAxisLabel* and neither a catAxisTitle nor the catGridLine (#479), and the
// line, tick, position, rotation and placement options are not measured there. bar3D draws its category labels along
// the floor's front edge, and a scatter or bubble chart's x axis is a value axis, which PptxGenJS writes from the
// catAxis* options: neither places its axes elsewhere, skips labels or formats categories.
const CATEGORY_AXIS_UNDRAWN_OPTIONS: Readonly<Partial<Record<NormalizedChartType, ReadonlySet<string>>>> = Object.freeze({
    radar: new Set<string>([
        ...CATEGORY_AXIS_PLACEMENT_OPTIONS,
        "catGridLine",
        "catAxisHidden",
        "catAxisLabelPos",
        "catAxisLabelRotate",
        "catAxisLineColor",
        "catAxisLineShow",
        "catAxisLineSize",
        "catAxisLineStyle",
        "catAxisMajorTickMark",
        "catAxisMinorTickMark",
    ]),
    bar3D: new Set<string>(CATEGORY_AXIS_PLACEMENT_OPTIONS),
    scatter: new Set<string>(CATEGORY_AXIS_PLACEMENT_OPTIONS),
    bubble: new Set<string>(CATEGORY_AXIS_PLACEMENT_OPTIONS),
});
// The chart types whose data table is drawn: a vertical category axis's columns hold it. LibreOffice 26.2 draws a
// horizontal bar chart's data table out of line with its categories (#567).
const DATA_TABLE_TYPES = new Set<NormalizedChartType>(["area", "bar", "line", "mixed"]);
const DATA_BORDER_TYPES = new Set<NormalizedChartType>(["area", "bar", "bar3D", "bubble", "doughnut", "mixed", "pie"]);
const TICK_MARKS = new Set(["none", "inside", "outside", "cross"]);
// bar3D draws its value axis on a wall, where LibreOffice 26.2.6 draws neither an axis line nor tick marks whatever
// valAxisLine* and valAxis*TickMark say (#561); it draws no logarithmic or reversed 3D axis or horizontal 3D bars yet.
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
function valueAxisDataRange(type: NormalizedChartType, series: readonly NormalizedChartSeries[]): [number, number] {
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

function finiteNumber(value: number | undefined, name: string, fallback: number): number {
    const normalized = value ?? fallback;

    if (!Number.isFinite(normalized)) {
        throw new TypeError(`${name} must be a finite number`);
    }

    return normalized;
}

/** How a value axis is scaled: LibreOffice's automatic 2D scale, its bar3D scale, or its radar scale in at most `maxIncrements` steps. */
type ValueAxisScaleKind = { kind: "linear" } | { kind: "bar3D" } | { kind: "radar"; maxIncrements: number };

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
    // A bar3D or radar axis takes LibreOffice's 3D scale over its data with 0 included; renderBar3DSvg takes a longer
    // step where the drawn axis is short, and a radar axis takes at most the target's radarMaxIncrements steps.
    const bar3D = scaleKind.kind === "bar3D";
    const bar3DLow = Math.min(0, dataRange[0]);
    const bar3DHigh = Math.max(0, dataRange[1]);
    const fixedScale = { minimum: fixedMinimum, maximum: fixedMaximum, majorUnit };
    const scale =
        scaleKind.kind === "bar3D"
            ? bar3DAxisScale(bar3DLow, bar3DHigh, fixedScale, PPTX_DEFAULTS.chart.automaticScale.maxIncrements)
            : scaleKind.kind === "radar"
              ? bar3DAxisScale(bar3DLow, bar3DHigh, fixedScale, scaleKind.maxIncrements)
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

    if (axisOptions?.valAxisMajorTickMark === "cross" || axisOptions?.valAxisMinorTickMark === "cross") {
        const option = axisOptions.valAxisMajorTickMark === "cross" ? "valAxisMajorTickMark" : "valAxisMinorTickMark";

        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: [`${path}${option}[cross]`] });
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
        majorTickMark: axisOptions?.valAxisMajorTickMark ?? defaults.majorTickMark,
        minorTickMark: axisOptions?.valAxisMinorTickMark ?? defaults.minorTickMark,
        gridLineColor: normalizeColor(gridLine.color ?? PPTX_DEFAULTS.chart.gridLine.color),
        gridLineWidth: positiveNumber(gridLine.size, "valGridLine.size", PPTX_DEFAULTS.chart.gridLine.widthPt),
        gridLineStyle: gridLine.style ?? PPTX_DEFAULTS.chart.gridLine.style,
        gridLineCap: gridLine.cap ?? PPTX_DEFAULTS.chart.gridLine.cap,
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

    const showSeriesName = !pieLike && (options.showSerName ?? PPTX_DEFAULTS.chart.showSeriesName);

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

export function normalizeChart(type: ChartInputType, data: readonly PptxGenJS.OptsChartData[], options: PptxGenJS.IChartOpts = {}, target: RenderingTarget = DEFAULT_RENDERING_TARGET): NormalizedChart {
    const profile = RENDERING_TARGETS[target];
    const mixed = Array.isArray(type);
    const chartTypes = mixed ? type.map(chart => chart.type) : [type];

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
                if (chart.options[key as keyof PptxGenJS.IChartOpts] !== undefined && !SUPPORTED_MIXED_SERIES_OPTIONS.has(key)) {
                    unsupported.push(`series[${index}].options.${key}`);
                }
            }
        }

        if (unsupported.length) {
            throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: unsupported.sort() });
        }
    }

    const requestedValueAxes = options.valAxes ?? [];
    const unsupportedAxisOptions: string[] = [];

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

    // PptxGenJS writes no axes on a pie or doughnut, so the category-axis options draw nothing there. A minMax
    // orientation and an autoZero crossing are the defaults, drawn as they are.
    const undrawnCategoryOptions = mixed ? undefined : CATEGORY_AXIS_UNDRAWN_OPTIONS[type as NormalizedChartType];

    for (const key of undrawnCategoryOptions ?? []) {
        const value: unknown = options[key as keyof PptxGenJS.IChartOpts];

        if (value !== undefined && value !== "minMax" && value !== "autoZero") {
            unsupportedAxisOptions.push(`${key}[${String(type)}]`);
        }
    }

    // bar3D and bubble charts draw their category or x labels below the plot area, or none.
    if ((type === "bar3D" || type === "bubble") && options.catAxisLabelPos === "high") {
        unsupportedAxisOptions.push(`catAxisLabelPos[high][${type}]`);
    }

    // bar3D draws its data labels without a box, and a radar chart its value labels with a formatter of its own.
    if (type === "bar3D" && options.dataLabelBkgrdColors) {
        unsupportedAxisOptions.push("dataLabelBkgrdColors[bar3D]");
    }

    if (type === "radar" && options.valAxisDisplayUnit !== undefined) {
        unsupportedAxisOptions.push("valAxisDisplayUnit[radar]");
    }

    if (options.showDataTable && !mixed && !DATA_TABLE_TYPES.has(type as NormalizedChartType)) {
        unsupportedAxisOptions.push(`showDataTable[${String(type)}]`);
    }

    if (options.showDataTable && options.barDir === "bar") {
        unsupportedAxisOptions.push("showDataTable[barDir bar]");
    }

    // A secondary value axis has a category axis of its own, which PptxGenJS crosses and orients by the same options.
    const crossesElsewhere = (options.catAxisCrossesAt ?? "autoZero") !== "autoZero" || (options.valAxisCrossesAt ?? "autoZero") !== "autoZero";

    if (requestedValueAxes.length > 1 && (crossesElsewhere || (options.catAxisOrientation ?? "minMax") !== "minMax")) {
        unsupportedAxisOptions.push("valAxes[secondary][catAxisCrossesAt, valAxisCrossesAt or catAxisOrientation]");
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

    if (!mixed && PLOT_AREA_UNDRAWN_TYPES.has(type as NormalizedChartType)) {
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
    const pointColor = (colors: readonly string[], value: number, index: number): CssColor => {
        const source = value < 0 && invertedColors ? invertedColors : colors;

        return normalizeColor(source[index % source.length] ?? "");
    };
    const withPointStyle = (item: NormalizedChartSeries, colors: readonly string[]): NormalizedChartSeries => {
        if (item.type === "bar" && options.lineSize !== 0) {
            return { ...item, varyColors: true, ...(invertedColors ? { pointFills: item.values.map((value, index) => pointColor(colors, value, index)) } : {}) };
        }

        // shortcut: the automatic fill of the first series; a single-series bar group after other series in a combo
        // chart may take a later accent, measure it when a deck needs it.
        const automaticPointFill = normalizeColor(PPTX_DEFAULTS.theme.colors.accent1);

        if (item.type === "bar3D" && options.lineSize !== 0) {
            return { ...item, automaticPointFill, pointLineColors: item.values.map((value, index) => pointColor(colors, value, index)) };
        }

        return { ...item, automaticPointFill };
    };
    let series: NormalizedChartSeries[];
    // A combo chart series' chart type options, by series index.
    const seriesOptions: PptxGenJS.IChartOpts[] = [];

    if (!Array.isArray(type)) {
        series = normalizeSeries(normalizedType, data);

        if ((normalizedType === "bar" || normalizedType === "bar3D") && pointColors && series.length === 1 && series[0]) {
            series[0] = withPointStyle(series[0], pointColors);
        }
    } else {
        series = [];

        for (const chart of type) {
            for (const item of normalizeSeries(chart.type as NormalizedChartType, chart.data)) {
                // PptxGenJS restarts colors within each mixed chart-type group.
                const placed = { ...item, valueAxisIndex: chart.options?.secondaryValAxis ? (1 as const) : (0 as const) };

                series.push(chart.type === "bar" && chart.data.length === 1 && pointColors ? withPointStyle(placed, pointColors) : placed);
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
    const barStack = mixed && normalizedGrouping === "percentStacked" ? profile.comboPercentBarStack : normalizedGrouping === "stacked" || normalizedGrouping === "percentStacked" ? normalizedGrouping : undefined;

    for (const [index, item] of series.entries()) {
        const stack = item.type === "line" ? undefined : item.type === "area" ? (normalizedGrouping === "stacked" ? normalizedGrouping : undefined) : barStack;

        if (stack) {
            series[index] = { ...item, stack };
        }
    }

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
            series[index] = { ...item, dataLabels: normalizeDataLabels(item.type, { ...options, ...seriesOptions[index] }, normalizedGrouping, chartTypes, profile) };
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
    const tickMarks: string[] = [];

    for (const key of ["catAxisMajorTickMark", "catAxisMinorTickMark"] as const) {
        const mark: unknown = options[key];

        if (mark !== undefined && !TICK_MARKS.has(String(mark))) {
            throw new TypeError(`${key} must be "none", "inside", "outside" or "cross"; got ${JSON.stringify(mark)}`);
        }

        if (mark === "cross") {
            tickMarks.push(`${key}[cross]`);
        }
    }

    if (tickMarks.length) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: tickMarks });
    }

    const categoryCount = series[0]?.labels.length ?? 0;
    // `<c:crossesAt val="${catAxisCrossesAt}"/>` on the value axis is the 1-based category it crosses before; a string
    // is written `<c:crosses val="...">` ("autoZero", "min" or "max").
    const catCrossesAt: unknown = options.catAxisCrossesAt;
    let valueAxisCrossing: NormalizedChart["categoryAxis"]["valueAxisCrossing"];

    if (typeof catCrossesAt === "number") {
        if (!Number.isFinite(catCrossesAt)) {
            throw new TypeError(`catAxisCrossesAt must be a finite category number or "autoZero"; got ${catCrossesAt}`);
        }

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

        labelSkip = profile.readsCategoryLabelSkip && skip > 1 ? skip : undefined;
    }

    const labelFormatCode = options.catLabelFormatCode;

    if (labelFormatCode !== undefined && !isSupportedChartNumberFormat(labelFormatCode)) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["catLabelFormatCode[format]"] });
    }

    const orientation: string | undefined = options.catAxisOrientation;

    if (orientation !== undefined && orientation !== "minMax" && orientation !== "maxMin") {
        throw new TypeError(`catAxisOrientation must be "minMax" or "maxMin"; got ${JSON.stringify(orientation)}`);
    }

    const categoryAxis: NormalizedChart["categoryAxis"] = {
        showTitle: options.showCatAxisTitle ?? false,
        title: options.catAxisTitle || catDefaults.title,
        titleColor: normalizeColor(options.catAxisTitleColor || catDefaults.titleColor),
        titleFontFace: options.catAxisTitleFontFace || PPTX_DEFAULTS.chart.fontFace,
        titleFontSize: positiveNumber(options.catAxisTitleFontSize, "catAxisTitleFontSize", catDefaults.titleFontSizePt),
        titleRotate: options.catAxisTitleRotate === undefined ? undefined : finiteNumber(options.catAxisTitleRotate, "catAxisTitleRotate", 0),
        labelColor: normalizeColor(options.catAxisLabelColor || catDefaults.labelColor),
        labelFontFace: options.catAxisLabelFontFace || PPTX_DEFAULTS.chart.fontFace,
        labelFontSize: positiveNumber(options.catAxisLabelFontSize, "catAxisLabelFontSize", catDefaults.labelFontSizePt),
        labelBold: options.catAxisLabelFontBold ?? catDefaults.labelBold,
        labelItalic: options.catAxisLabelFontItalic ?? catDefaults.labelItalic,
        hidden: options.catAxisHidden ?? catDefaults.hidden,
        // PptxGenJS writes `nextTo` on a scatter chart's x axis whatever catAxisLabelPos says.
        labelPosition: normalizedType === "scatter" ? "nextTo" : (options.catAxisLabelPos ?? (options.barDir === "bar" ? catDefaults.barLabelPosition : catDefaults.columnLabelPosition)),
        labelRotate: finiteNumber(options.catAxisLabelRotate, "catAxisLabelRotate", catDefaults.labelRotate),
        ...(labelSkip ? { labelSkip } : {}),
        ...(labelFormatCode !== undefined ? { labelFormatCode } : {}),
        ...(orientation === "maxMin" ? { reversed: true } : {}),
        lineColor: normalizeColor(options.catAxisLineColor || catDefaults.lineColor),
        lineWidth: positiveNumber(options.catAxisLineSize || undefined, "catAxisLineSize", catDefaults.lineWidthPt),
        lineStyle: options.catAxisLineStyle ?? catDefaults.lineStyle,
        // PptxGenJS writes `<c:minorTickMark val="none"/>` on a scatter chart's axes, as its major ones.
        minorTickMark: normalizedType === "scatter" ? "none" : (options.catAxisMinorTickMark ?? catDefaults.minorTickMark),
        ...(valueAxisCrossing ? { valueAxisCrossing } : {}),
        gridLineColor: normalizeColor(catGridLine?.color ?? PPTX_DEFAULTS.chart.gridLine.color),
        gridLineWidth: positiveNumber(catGridLine?.size, "catGridLine.size", PPTX_DEFAULTS.chart.gridLine.widthPt),
        gridLineStyle: catGridLine ? (catGridLine.style ?? PPTX_DEFAULTS.chart.gridLine.style) : catDefaults.gridLineStyle,
        gridLineCap: catGridLine?.cap ?? PPTX_DEFAULTS.chart.gridLine.cap,
    };
    const primaryScale: ValueAxisScaleKind = normalizedType === "bar3D" ? { kind: "bar3D" } : normalizedType === "radar" ? { kind: "radar", maxIncrements: profile.radarMaxIncrements } : { kind: "linear" };
    const primaryAxis = normalizeValueAxis(valueAxisOptions[0] ?? options, primaryDataRange, primaryScale, valueAxisOptions[0] ? "valAxes[0]." : "", chartTypes, normalizedType === "scatter" ? scatterGridLine : {}, profile.readsValueAxisDisplayUnits);
    const valueAxes = [categoryAxisCrossing ? { ...primaryAxis, categoryAxisCrossing } : primaryAxis];

    if (normalizedType === "bar3D" && primaryAxis.displayUnit?.label) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["valAxisDisplayUnitLabel[bar3D]"] });
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

    // PptxGenJS writes the depth axis (makeSerAxis) on a bar3D chart only.
    const seriesAxis: NormalizedChart["seriesAxis"] =
        normalizedType === "bar3D"
            ? {
                  hidden: options.serAxisHidden ?? false,
                  reversed: serOrientation === "maxMin",
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
    const dataTable: NormalizedChart["dataTable"] = options.showDataTable
        ? {
              horizontalBorder: profile.drawsDataTableBorders && options.showDataTableHorzBorder === true,
              verticalBorder: profile.drawsDataTableBorders && options.showDataTableVertBorder === true,
              outline: profile.drawsDataTableBorders && options.showDataTableOutline === true,
              keys: profile.drawsDataTableBorders && options.showDataTableKeys === true,
              // shortcut: the theme's minor font (+mn-lt) is the default theme's; a presentation theme's body face is
              // not passed to charts yet, pass it when a deck sets one with a data table.
              fontFace: PPTX_DEFAULTS.theme.bodyFontFace,
              fontSize: checkedSize(options.dataTableFontSize || tableDefaults.fontSizePt, "dataTableFontSize", PPTX_DEFAULTS.limits.fontSizePt),
              color: normalizeColor(tableDefaults.color),
              borderColor: normalizeColor(tableDefaults.borderColor),
              borderWidth: tableDefaults.borderWidthPt,
          }
        : undefined;
    // dataBorder: `<a:ln w="${pt || 0.75}">` in its color on bar, bar3D and area series, bubbles and pie slices; a line,
    // radar or scatter series has its own line.
    const dataBorder: NormalizedChart["dataBorder"] =
        options.dataBorder && DATA_BORDER_TYPES.has(normalizedType)
            ? {
                  color: normalizeColor(options.dataBorder.color || PPTX_DEFAULTS.chart.dataBorder.color),
                  width: checkedSize(options.dataBorder.pt || PPTX_DEFAULTS.chart.dataBorder.widthPt, "dataBorder.pt", PPTX_DEFAULTS.limits.lineWidthPt),
              }
            : undefined;
    // `<c:overlap val="${stacked ? 100 : barOverlapPct || 0}"/>` on a bar chart; a stack is one bar already.
    const barOverlapPercent = (normalizedType === "bar" || normalizedType === "mixed") && !barStack ? Math.max(-100, Math.min(100, finiteNumber(options.barOverlapPct, "barOverlapPct", 0))) : 0;

    if (secondaryDataRange) {
        valueAxes.push(normalizeValueAxis(valueAxisOptions[1], secondaryDataRange, { kind: "linear" }, "valAxes[1].", chartTypes, {}, profile.readsValueAxisDisplayUnits));
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
        colors: palette.map(color => normalizeColor(color)),
        colorOpacity: Math.max(0, Math.min(1, (options.chartColorsOpacity ?? 100) / 100)),
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
        lineSmooth: options.lineSmooth ?? PPTX_DEFAULTS.chart.lineSmooth,
        lineSymbol: options.lineDataSymbol ?? PPTX_DEFAULTS.chart.lineDataSymbol,
        lineSymbolSize: checkedSize(options.lineDataSymbolSize ?? PPTX_DEFAULTS.chart.lineDataSymbolSize, "lineDataSymbolSize", PPTX_DEFAULTS.limits.chartMarkerSizePt),
        // PptxGenJS: `<a:ln w="${lineDataSymbolLineSize || 0.75}"><a:solidFill>${lineDataSymbolLineColor || seriesColor}`.
        ...(options.lineDataSymbolLineColor ? { lineSymbolLineColor: normalizeColor(options.lineDataSymbolLineColor) } : {}),
        lineSymbolLineSize: checkedSize(options.lineDataSymbolLineSize || PPTX_DEFAULTS.chart.lineDataSymbolLineSizePt, "lineDataSymbolLineSize", PPTX_DEFAULTS.limits.lineWidthPt),
        lineSize: checkedSize(options.lineSize ?? PPTX_DEFAULTS.chart.lineSizePt, "lineSize", PPTX_DEFAULTS.limits.lineWidthPt),
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
        categoryMajorTickMark: normalizedType === "scatter" ? "none" : (options.catAxisMajorTickMark ?? PPTX_DEFAULTS.chart.categoryMajorTickMark),
        valueAxisMinimum: primaryAxis.minimum,
        valueAxisMaximum: primaryAxis.maximum,
        valueAxes: normalizedType === "scatter" ? valueAxes.map(axis => ({ ...axis, majorTickMark: "none" as const, minorTickMark: "none" as const })) : valueAxes,
        categoryAxis,
        ...(seriesAxis ? { seriesAxis } : {}),
        ...(dataTable ? { dataTable } : {}),
        ...(dataBorder ? { dataBorder } : {}),
    };
}
