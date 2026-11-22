import type PptxGenJS from "pptxgenjs";
import { automaticAxisScale, type FixedAxisBound } from "#src/chart/axisScale";
import { UnsupportedChartError } from "#src/chart/errors";
import { isSupportedChartNumberFormat } from "#src/chart/format";
import { PPTX_DEFAULTS } from "#src/defaults";
import type { NormalizedChart, NormalizedChartSeries, NormalizedChartType } from "@pdflapas/model";
import { checkedSize } from "./size";
import { normalizeColor, normalizeFill } from "./style";

const SUPPORTED_TYPES = new Set<NormalizedChartType>(["area", "bar", "bar3D", "bubble", "doughnut", "line", "mixed", "pie", "radar", "scatter"]);

const MIXED_CHART_TYPES = new Set<PptxGenJS.CHART_NAME>(["area", "bar", "line"]);

export const CATEGORY_AXIS_OPTIONS = [
    "catAxisLabelColor",
    "catAxisLabelFontBold",
    "catAxisLabelFontFace",
    "catAxisLabelFontItalic",
    "catAxisLabelFontSize",
    "catAxisTitle",
    "catAxisTitleColor",
    "catAxisTitleFontFace",
    "catAxisTitleFontSize",
    "catAxisTitleRotate",
    "catGridLine",
    "showCatAxisTitle",
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
    "bar3DShape",
    "barDir",
    "barGrouping",
    "barGapDepthPct",
    "barGapWidthPct",
    "v3DPerspective",
    "v3DRAngAx",
    "v3DRotX",
    "v3DRotY",
    "showValAxisTitle",
    "valAxes",
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
    "dataLabelPosition",
    "dataLabelFormatCode",
    "dataLabelColor",
    "dataLabelFontBold",
    "dataLabelFontFace",
    "dataLabelFontItalic",
    "dataLabelFontSize",
    ...CATEGORY_AXIS_OPTIONS,
    "layout",
    "plotArea",
] as const);

// The plot rectangle is the ECharts grid, which pie, doughnut, radar and bar3D charts do not use.
const PLOT_AREA_UNDRAWN_TYPES = new Set<NormalizedChartType>(["bar3D", "doughnut", "pie", "radar"]);
const SUPPORTED_PLOT_AREA_OPTIONS: Readonly<Record<string, ReadonlySet<string>>> = Object.freeze({
    // genXmlColorSelection reads type, color, transparency and the deprecated alpha.
    fill: new Set(["alpha", "color", "transparency", "type"]),
    // PptxGenJS writes pt and color and ignores type: the border is always a solid line.
    border: new Set(["color", "pt", "type"]),
});
const LAYOUT_KEYS = ["x", "y", "w", "h"] as const;

const SUPPORTED_CHART_OPTION_SET = new Set<string>(SUPPORTED_CHART_OPTIONS);
const SUPPORTED_MIXED_SERIES_OPTIONS = new Set(["secondaryValAxis"]);
const SUPPORTED_VALUE_AXIS_OPTIONS = new Set([
    "showValAxisTitle",
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
const CATEGORY_AXIS_UNDRAWN_TYPES = new Set<NormalizedChartType>(["bar3D", "bubble", "radar", "scatter"]);
const SUPPORTED_GRID_LINE_OPTIONS = new Set(["cap", "color", "size", "style"]);
const BAR_3D_SHAPES = new Set(["box", "cylinder", "cone", "coneToMax", "pyramid", "pyramidToMax"]);
const DATA_LABEL_CHART_TYPES = new Set<NormalizedChartType>(["bar", "doughnut", "line", "pie"]);
// The dataLabelPosition values PptxGenJS 4.0.1 keeps; its option cleanup deletes any other (every one on a doughnut and,
// a quirk, `outEnd` on a clustered bar), so the chart type's default position applies.
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

    // Each stack's positive and negative totals per category, or their shares of the category's absolute total.
    for (const stack of stacks.values()) {
        const percent = stack[0]?.stack === "percentStacked";
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

            if (!percent) {
                values.push(positive, negative);
            } else if (total === 0) {
                values.push(0, 0);
            } else {
                values.push((positive / total) * 100, (negative / total) * 100);
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

function normalizeValueAxis(axisOptions: PptxGenJS.IChartPropsAxisVal | undefined, dataRange: [number, number], path: string, chartTypes: PptxGenJS.CHART_NAME[], defaultGridLine: PptxGenJS.OptsChartGridLine): NormalizedChart["valueAxes"][number] {
    const defaults = PPTX_DEFAULTS.chart.valueAxis;
    // Check the axis's bounds, scale, unit, tick marks and label format. An automatic bound scales from a fixed one (#556).
    const fixedMinimum = axisOptions?.valAxisMinVal;
    const fixedMaximum = axisOptions?.valAxisMaxVal;
    const fixed: FixedAxisBound = fixedMinimum !== undefined ? { kind: "minimum", value: fixedMinimum } : fixedMaximum !== undefined ? { kind: "maximum", value: fixedMaximum } : { kind: "none" };
    const scale = automaticAxisScale(dataRange[0], dataRange[1], fixed, true);
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

    const majorUnit = axisOptions?.valAxisMajorUnit;

    if (majorUnit !== undefined && (!Number.isFinite(majorUnit) || majorUnit <= 0)) {
        throw new TypeError("valAxisMajorUnit must be a positive finite number");
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

    return {
        minimum,
        maximum,
        majorUnit,
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

type ChartInputType = PptxGenJS.CHART_NAME | PptxGenJS.IChartMulti[];

export function normalizeChart(type: ChartInputType, data: readonly PptxGenJS.OptsChartData[], options: PptxGenJS.IChartOpts = {}): NormalizedChart {
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

    // PptxGenJS writes no axes on a pie or doughnut, so the category-axis options draw nothing there; on the other
    // chart types without a 2D category axis (bar3D, radar, scatter, bubble) they are not drawn yet.
    if (!mixed && CATEGORY_AXIS_UNDRAWN_TYPES.has(type as NormalizedChartType)) {
        for (const key of CATEGORY_AXIS_OPTIONS) {
            if (options[key] !== undefined) {
                unsupportedAxisOptions.push(`${key}[${String(type)}]`);
            }
        }
    }

    for (const [key, value] of Object.entries(options.plotArea ?? {})) {
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
        for (const key of ["layout", "plotArea"] as const) {
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
    const pointColors = options.chartColors && options.chartColors.length > 1 ? options.chartColors : undefined;
    const withPointStyle = (item: NormalizedChartSeries, colors: readonly string[]): NormalizedChartSeries => {
        if (item.type === "bar" && options.lineSize !== 0) {
            return { ...item, varyColors: true };
        }

        // shortcut: the automatic fill of the first series; a single-series bar group after other series in a combo
        // chart may take a later accent, measure it when a deck needs it.
        const automaticPointFill = normalizeColor(PPTX_DEFAULTS.theme.colors.accent1);

        if (item.type === "bar3D" && options.lineSize !== 0) {
            return { ...item, automaticPointFill, pointLineColors: item.values.map((_, index) => normalizeColor(colors[index % colors.length] ?? "")) };
        }

        return { ...item, automaticPointFill };
    };
    let series: NormalizedChartSeries[];

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
    const barStack = normalizedGrouping === "stacked" || normalizedGrouping === "percentStacked" ? normalizedGrouping : undefined;

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

    // The data labels. PptxGenJS writes showCatName and showPercent only on pie and doughnut points (and custom labels
    // for showLabel on scatter), so on a bar or line chart showLabel and showPercent draw nothing, as in PowerPoint.
    const pieLike = normalizedType === "pie" || normalizedType === "doughnut";
    const drawingLabelOptions: string[] = [];

    if (options.showValue) {
        drawingLabelOptions.push("showValue");
    }

    if (options.showLabel && (pieLike || normalizedType === "scatter")) {
        drawingLabelOptions.push("showLabel");
    }

    if (options.showPercent && pieLike) {
        drawingLabelOptions.push("showPercent");
    }

    if (drawingLabelOptions.length && !DATA_LABEL_CHART_TYPES.has(normalizedType)) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: drawingLabelOptions.map(option => `${option}[${normalizedType}]`) });
    }

    const stackedBars = normalizedGrouping === "stacked" || normalizedGrouping === "percentStacked";
    const defaultPositions = PPTX_DEFAULTS.chart.dataLabelPosition;
    let labelPosition: NormalizedChart["dataLabels"]["position"] =
        normalizedType === "line" ? defaultPositions.line : normalizedType === "pie" ? defaultPositions.pie : normalizedType === "doughnut" ? defaultPositions.doughnut : stackedBars ? defaultPositions.stacked : defaultPositions.clustered;
    const keptPositions = normalizedType === "line" ? LINE_LABEL_POSITIONS : normalizedType === "pie" ? PIE_LABEL_POSITIONS : normalizedType !== "bar" ? NO_LABEL_POSITIONS : stackedBars ? STACKED_BAR_LABEL_POSITIONS : CLUSTERED_BAR_LABEL_POSITIONS;

    if (options.dataLabelPosition !== undefined && keptPositions.has(options.dataLabelPosition)) {
        labelPosition = options.dataLabelPosition;
    }

    if (normalizedType === "pie" && (labelPosition === "inEnd" || labelPosition === "outEnd") && drawingLabelOptions.length) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: [`dataLabelPosition[${labelPosition}]`] });
    }

    // PptxGenJS: `dataLabelFormatCode || (pie or doughnut: showPercent ? '0%' : 'General'; scatter: 'General') || '#,##0'`.
    // On a pie LibreOffice applies it to the percentage when showPercent is set and draws the value in General (renders
    // of PptxGenJS pies: `0.0` with showPercent draws 0.5; showValue with showPercent draws 50.25 and 50%).
    const pieFormatCode = options.showPercent ? PPTX_DEFAULTS.chart.piePercentFormatCode : PPTX_DEFAULTS.chart.pieDataLabelFormatCode;
    const labelFormatCode = options.dataLabelFormatCode || (pieLike ? pieFormatCode : normalizedType === "scatter" ? PPTX_DEFAULTS.chart.scatterDataLabelFormatCode : PPTX_DEFAULTS.chart.dataLabelFormatCode);

    if (!isSupportedChartNumberFormat(labelFormatCode)) {
        throw new UnsupportedChartError({ reason: "chart-option", chartTypes, unsupportedOptions: ["dataLabelFormatCode[format]"] });
    }

    const dataLabels: NormalizedChart["dataLabels"] = {
        showValue: options.showValue ?? PPTX_DEFAULTS.chart.showValue,
        showCategory: pieLike && (options.showLabel ?? PPTX_DEFAULTS.chart.showLabel),
        showPercent: pieLike && (options.showPercent ?? PPTX_DEFAULTS.chart.showPercent),
        position: labelPosition,
        valueFormat: pieLike && options.showPercent ? PPTX_DEFAULTS.chart.pieDataLabelFormatCode : labelFormatCode,
        percentFormat: labelFormatCode,
        color: normalizeColor(options.dataLabelColor || PPTX_DEFAULTS.chart.dataLabelColor),
        fontFace: options.dataLabelFontFace || PPTX_DEFAULTS.chart.fontFace,
        fontSize: checkedSize(options.dataLabelFontSize || PPTX_DEFAULTS.chart.dataLabelFontSizePt, "dataLabelFontSize", PPTX_DEFAULTS.limits.fontSizePt),
        bold: options.dataLabelFontBold ?? false,
        italic: options.dataLabelFontItalic ?? false,
    };
    const axisSeries = (axisIndex: 0 | 1) => series.filter(item => item.valueAxisIndex === axisIndex);
    const primaryDataRange = valueAxisDataRange(normalizedType, axisSeries(0));
    const secondaryDataRange = hasSecondaryValueAxis ? valueAxisDataRange(normalizedType, axisSeries(1)) : undefined;
    const catDefaults = PPTX_DEFAULTS.chart.categoryAxis;
    // PptxGenJS 4.0.1 gives a scatter chart (not a bubble chart) `{ color: 'D9D9D9', size: 1 }` gridlines on both axes
    // when catGridLine or valGridLine is unset.
    const scatterGridLine: PptxGenJS.OptsChartGridLine = { color: PPTX_DEFAULTS.chart.scatterGridLine.color, size: PPTX_DEFAULTS.chart.scatterGridLine.widthPt };
    const catGridLine = options.catGridLine ?? (normalizedType === "scatter" ? scatterGridLine : undefined);
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
        gridLineColor: normalizeColor(catGridLine?.color ?? PPTX_DEFAULTS.chart.gridLine.color),
        gridLineWidth: positiveNumber(catGridLine?.size, "catGridLine.size", PPTX_DEFAULTS.chart.gridLine.widthPt),
        gridLineStyle: catGridLine ? (catGridLine.style ?? PPTX_DEFAULTS.chart.gridLine.style) : catDefaults.gridLineStyle,
        gridLineCap: catGridLine?.cap ?? PPTX_DEFAULTS.chart.gridLine.cap,
    };
    const primaryAxis = normalizeValueAxis(valueAxisOptions[0] ?? options, primaryDataRange, valueAxisOptions[0] ? "valAxes[0]." : "", chartTypes, normalizedType === "scatter" ? scatterGridLine : {});
    const valueAxes = [primaryAxis];

    if (secondaryDataRange) {
        valueAxes.push(normalizeValueAxis(valueAxisOptions[1], secondaryDataRange, "valAxes[1].", chartTypes, {}));
    }

    // PptxGenJS 4.0.1: a fill without a color writes <a:noFill/>; a border defaults a missing or zero pt and a missing
    // color to DEF_CHART_BORDER (1pt 363636). The deprecated `fill` and `border` options stay rejected above.
    const plotFill = options.plotArea?.fill;
    const plotBorder = options.plotArea?.border;
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
        // PptxGenJS never applies fontFace/fontSize to the legend.
        legendFontFace: options.legendFontFace ?? PPTX_DEFAULTS.theme.bodyFontFace,
        legendFontSize: checkedSize(options.legendFontSize ?? PPTX_DEFAULTS.chart.legendFontSizePt, "legendFontSize", PPTX_DEFAULTS.limits.fontSizePt),
        legendColor: normalizeColor(options.legendColor ?? PPTX_DEFAULTS.text.color),
        // Bar options.
        barDirection: options.barDir === "bar" ? ("bar" as const) : PPTX_DEFAULTS.chart.barDirection,
        // PptxGenJS 4.0.1 sets a stacked grouping's `barGapWidthPct` to 50 when it is unset or 0 (`if (!barGapWidthPct)`).
        barGapWidthPercent: Math.max(0, Math.min(1000, (stackedBars && !options.barGapWidthPct ? PPTX_DEFAULTS.chart.stackedGapWidthPct : options.barGapWidthPct) ?? PPTX_DEFAULTS.chart.barGapWidthPct)),
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
        altText: options.altText ?? "",
        grouping: normalizedGrouping,
        gridLineColor: normalizeColor(PPTX_DEFAULTS.chart.gridLine.color),
        gridLineWidth: PPTX_DEFAULTS.chart.gridLine.widthPt,
        axisLineVisible: PPTX_DEFAULTS.chart.categoryAxisLineVisible,
        // PptxGenJS 4.0.1 writes `<c:majorTickMark val="none"/>` and `<c:minorTickMark val="none"/>` on both scatter
        // axes, whatever catAxisMajorTickMark or valAxisMajorTickMark say.
        categoryMajorTickMark: normalizedType === "scatter" ? "none" : PPTX_DEFAULTS.chart.categoryMajorTickMark,
        valueAxisMinimum: primaryAxis.minimum,
        valueAxisMaximum: primaryAxis.maximum,
        valueAxes: normalizedType === "scatter" ? valueAxes.map(axis => ({ ...axis, majorTickMark: "none" as const, minorTickMark: "none" as const })) : valueAxes,
        categoryAxis,
    };
}
