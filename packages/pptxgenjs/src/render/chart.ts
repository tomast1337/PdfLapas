import { BarChart, LineChart, PieChart, RadarChart, ScatterChart } from "echarts/charts";
import { GraphicComponent, GridComponent, LegendComponent, RadarComponent, TitleComponent } from "echarts/components";
import { type EChartsCoreOption, format, init, use } from "echarts/core";
import { SVGRenderer } from "echarts/renderers";
import { automaticAxisInterval, automaticAxisScale, type FixedAxisBound, paddedAxisScale } from "#src/chart/axisScale";
import { pointLabelText, showsPointLabel } from "#src/chart/dataLabel";
import { formatChartNumber } from "#src/chart/format";
import { PPTX_DEFAULTS } from "#src/defaults";
import { builtInFontMetrics } from "#src/fontMetrics";
import type { ChartExtensionInput } from "#src/chart/types";
import type { NormalizedChart, NormalizedChartValueAxis, NormalizedObjectStyle } from "@pdflapas/model";
import { pointsToPixels } from "@pdflapas/model";
import { renderBar3DSvg } from "./chart3d";
import { type InlineStyle, styleAttribute } from "./css";
import { attributeName, descendants, firstElementChild, getAttribute, parseMarkup, SVG_NAMESPACE, serializeOuter, setAttribute } from "./fragment";
import { element, type Html } from "./markup";
import { applyObjectStyle } from "./style";

use([BarChart, LineChart, PieChart, RadarChart, ScatterChart, GraphicComponent, GridComponent, LegendComponent, RadarComponent, TitleComponent, SVGRenderer]);

// LibreOffice draws OOXML's `dot` marker as its right-pointing arrow (oox maps XML_dot to chart2 StandardSymbol 4,
// Symbol_Right) and `dash` as its horizontal bar (Symbol_HorizontalBar), 0.2 of the marker size tall: measured in its
// PDF export of a synthetic line chart with 14pt markers, a 14pt triangle and a 14pt by 3pt bar.
// shortcut: PowerPoint draws `dot` as a small round dot and `dash` as its own short bar, not measured yet (#500).
const DOT_SYMBOL = "path://M0 0L2 1L0 2Z";
const DASH_SYMBOL_HEIGHT_RATIO = 0.2;

function lineSymbol(symbol: NormalizedChart["lineSymbol"]): string {
    if (symbol === "dot") {
        return DOT_SYMBOL;
    }

    if (symbol === "dash") {
        return "rect";
    }

    return symbol;
}

function lineType(style: "solid" | "dash" | "dot"): "solid" | "dashed" | "dotted" {
    if (style === "dash") {
        return "dashed";
    }

    if (style === "dot") {
        return "dotted";
    }

    return "solid";
}

/**
 * Tick length in px. shortcut: a cross tick (LibreOffice 26.2.6 draws it across the axis line) is drawn outside at
 * twice the length, as ECharts draws ticks on one side only; overlay an inside-tick axis if a fixture needs it.
 */
export const TICK_LENGTH = 5;
/** ECharts' default axisLabel.margin, the px between a tick and its label. */
const AXIS_LABEL_MARGIN = 8;
/** ECharts' default axisLabel.textMargin[1], the px it adds on each side of a label's measured width when fitting it. */
const AXIS_LABEL_TEXT_MARGIN = 3;
/** ECharts' default title.padding, the px between the chart's top edge and its title. */
const TITLE_PADDING = 5;
/** ECharts' default radar.radius ("50%"), the share of half the frame's shorter side the radar's circle takes. */
const RADAR_RADIUS_RATIO = 0.5;

function tickOption(mark: NormalizedChart["valueAxes"][number]["majorTickMark"], color: string, width: number) {
    return {
        show: mark !== "none",
        inside: mark === "inside",
        length: mark === "cross" ? TICK_LENGTH * 2 : TICK_LENGTH,
        lineStyle: { color, width: pointsToPixels(width) },
    };
}

/** An axis's gridlines (c:majorGridlines) as an ECharts splitLine. */
function gridLineOption(axis: Pick<NormalizedChart["categoryAxis"], "gridLineCap" | "gridLineColor" | "gridLineStyle" | "gridLineWidth">) {
    return {
        show: axis.gridLineStyle !== "none",
        lineStyle: {
            color: axis.gridLineColor,
            width: pointsToPixels(axis.gridLineWidth),
            type: axis.gridLineStyle === "none" ? "solid" : lineType(axis.gridLineStyle),
            cap: axis.gridLineCap === "flat" ? "butt" : axis.gridLineCap,
        },
    } as const;
}

/**
 * The ECharts value axis of a chart's value axis from `min` to `max`, on the left (index 0) or the right; a horizontal
 * bar chart's first axis sits at the bottom with its labels `labelMargin` px below it.
 */
function valueAxisOption(axis: NormalizedChartValueAxis, index: number, min: number, max: number, labelMargin: { kind: "margin"; px: number } | { kind: "default" }) {
    return {
        type: axis.logScaleBase === undefined ? ("value" as const) : ("log" as const),
        logBase: axis.logScaleBase,
        min,
        max,
        interval: axis.logScaleBase === undefined ? (axis.majorUnit ?? automaticAxisInterval(min, max)) : undefined,
        // valAxisOrientation "maxMin": the maximum at the start, and the crossing axis with it (Office crosses at the minimum).
        inverse: axis.reversed === true,
        position: index === 1 ? ("right" as const) : ("left" as const),
        show: !axis.hidden,
        // A display unit's name (`<c:dispUnitsLbl/>`) at the axis's far end, where the axis has no title.
        // shortcut: not measured against PowerPoint's PDF export, which may turn it up beside the axis; measure it (#566).
        name: axis.showTitle ? axis.title : axis.displayUnit?.label,
        nameLocation: axis.showTitle || !axis.displayUnit?.label ? ("middle" as const) : ("end" as const),
        nameGap: axis.showTitle || !axis.displayUnit?.label ? 40 : AXIS_LABEL_MARGIN,
        nameRotate: axis.titleRotate,
        nameTextStyle: {
            color: axis.titleColor,
            fontFamily: axis.titleFontFace,
            fontSize: pointsToPixels(axis.titleFontSize),
        },
        splitNumber: Math.min(10, Math.max(1, Number.isInteger(axis.maximum - axis.minimum) ? axis.maximum - axis.minimum : 5)),
        axisLabel: {
            show: axis.labelPosition !== "none",
            inside: axis.labelPosition === "high",
            rotate: axis.labelRotate,
            color: axis.labelColor,
            fontFamily: axis.labelFontFace,
            fontSize: pointsToPixels(axis.labelFontSize),
            ...(labelMargin.kind === "margin" ? { margin: labelMargin.px } : {}),
            fontWeight: axis.labelBold ? "bold" : "normal",
            fontStyle: axis.labelItalic ? "italic" : "normal",
            formatter: (value: number) => formatChartNumber(axis.displayUnit ? value / axis.displayUnit.divisor : value, axis.labelFormatCode),
        },
        axisLine: {
            show: axis.lineVisible,
            lineStyle: {
                color: axis.lineColor,
                width: pointsToPixels(axis.lineWidth),
                type: lineType(axis.lineStyle),
            },
        },
        axisTick: tickOption(axis.majorTickMark, axis.lineColor, axis.lineWidth),
        minorTick: {
            ...tickOption(axis.minorTickMark, axis.lineColor, axis.lineWidth),
            splitNumber: 5,
        },
        splitLine: index === 0 ? gridLineOption(axis) : { show: false },
    };
}

/** ECharts label positions for each OOXML `c:dLblPos`: a column's (vertical), a horizontal bar's, a line point's and a slice's. */
type LabelPositions = Readonly<Partial<Record<NormalizedChart["dataLabels"]["position"], string>>>;
const COLUMN_LABEL_POSITIONS: LabelPositions = { outEnd: "top", inEnd: "insideTop", ctr: "inside", inBase: "insideBottom" };
const BAR_LABEL_POSITIONS: LabelPositions = { outEnd: "right", inEnd: "insideRight", ctr: "inside", inBase: "insideLeft" };
const POINT_LABEL_POSITIONS: LabelPositions = { t: "top", b: "bottom", l: "left", r: "right", ctr: "inside" };
// A negative bar's: its end is at the bottom (left), its base at the zero line, as LibreOffice 26.2 draws them (measured
// on its PDF export of synthetic PptxGenJS 4.0.1 column and horizontal bar charts with values 5, -3, 7, -6).
// shortcut: PowerPoint is not measured; both targets take these (#559).
const NEGATIVE_COLUMN_LABEL_POSITIONS: LabelPositions = { outEnd: "bottom", inEnd: "insideBottom", ctr: "inside", inBase: "insideTop" };
const NEGATIVE_BAR_LABEL_POSITIONS: LabelPositions = { outEnd: "left", inEnd: "insideLeft", ctr: "inside", inBase: "insideRight" };

/** The ECharts `label` of a series; `text` gives each data point's label. */
function dataLabelOption(labels: NormalizedChart["dataLabels"], position: string, text: (dataIndex: number) => string) {
    return {
        show: true,
        position,
        formatter: (params: { dataIndex: number }) => text(params.dataIndex),
        color: labels.color,
        fontFamily: labels.fontFace,
        fontSize: pointsToPixels(labels.fontSize),
        fontWeight: labels.bold ? "bold" : "normal",
        fontStyle: labels.italic ? "italic" : "normal",
        // ECharts outlines a label drawn over a mark; Office draws the text alone.
        textBorderWidth: 0,
    };
}

/** ECharts' px between a legend key and its text (LegendView: `textX = itemWidth + 5`). */
export const LEGEND_TEXT_GAP = 5;

/** ECharts' legend itemWidth, which a line key keeps. */
const LEGEND_LINE_KEY_WIDTH = 25;

/** A filled square; the legend's symbolKeepAspect keeps it square when a line key makes the items wider. */
const LEGEND_SQUARE_ICON = "path://M0 0H1V1H0Z";

/** zrender's SVGPainter.renderToString; zrender is an echarts dependency, not ours, so it is typed here. */
type SvgPainter = {
    renderToString(options: { cssAnimation: boolean; cssEmphasis: boolean }): string;
};

/** zrender prefixes SVG ids with its instance number (zr<N>-), taken from a process-wide counter. */
const ZRENDER_ID_PREFIX = /^zr\d+-/;

function renderEChartsSvg(width: number, height: number, option: EChartsCoreOption): string {
    // zrender checks the process-global `document` when measuring text. The
    // visual oracle installs jsdom globally for PptxGenJS, but jsdom exposes a
    // canvas whose getContext is intentionally unimplemented. Hide only that
    // configurable jsdom global during synchronous SSR so zrender uses its
    // deterministic server-side width table instead of probing canvas.
    const globalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");
    const isJsdom = globalDocument?.configurable && String((globalThis as { document?: Document }).document?.defaultView?.navigator.userAgent).includes("jsdom");

    if (isJsdom) {
        Reflect.deleteProperty(globalThis, "document");
    }

    let instance: ReturnType<typeof init> | undefined;

    try {
        instance = init(null, null, { renderer: "svg", ssr: true, width, height });
        instance.setOption({ ...option, animation: false });
        // Not renderToSVGString(): it cannot turn off the emphasis :hover CSS, whose class names
        // (zr<N>-cls-<M>) come from process-wide counters and made toHtml() differ between runs (#66).
        // A printed chart has no hover. `as`: ssr with renderer "svg" always creates zrender's SVGPainter.
        const painter = instance.getZr().painter as unknown as SvgPainter;

        return painter.renderToString({ cssAnimation: false, cssEmphasis: false });
    } finally {
        instance?.dispose();

        if (isJsdom && globalDocument) {
            Object.defineProperty(globalThis, "document", globalDocument);
        }
    }
}

const RESOURCE_ATTRIBUTE = /^(?:href|src|xlink:href)$/i;
const EXTERNAL_URL = /^(?:https?:|file:|\/\/)/i;
const EXTERNAL_CSS_URL = /url\(\s*["']?(?:https?:|file:|\/\/)/i;

/** `idScope` is unique per object in the document (slide and object index), so chart ids stay apart and deterministic. */
function chartContainer(svgSource: string, style: NormalizedObjectStyle, altText: string, idScope: string): Html {
    const css: InlineStyle = new Map();
    const objectAttributes = applyObjectStyle(css, style);

    css.set("overflow", "hidden");

    // The chart's SVG root as markup: parsed as `template.innerHTML` parses it, checked for external references,
    // its ids prefixed with `idScope` and its size set to the object box. Only the root's descendants are
    // checked and renamed (it was `svg.querySelectorAll()`), so the bytes match the DOM-built HTML.
    const root = firstElementChild(parseMarkup(svgSource.trim()));

    if (root.kind === "none" || root.element.namespaceURI !== SVG_NAMESPACE || root.element.tagName !== "svg") {
        throw new TypeError("Chart renderer did not return an SVG root element");
    }

    const svg = root.element;

    const elements = Array.from(descendants(svg));

    for (const element of elements) {
        for (const attribute of element.attrs) {
            if ((RESOURCE_ATTRIBUTE.test(attributeName(attribute)) && EXTERNAL_URL.test(attribute.value)) || EXTERNAL_CSS_URL.test(attribute.value)) {
                throw new TypeError("Chart SVG must not reference external resources");
            }
        }
    }

    const ids = new Map<string, string>();

    for (const element of elements) {
        const id = getAttribute(element, "id");

        if (id.kind === "absent") {
            continue;
        }

        const oldId = id.value;

        const newId = `${idScope}-${oldId.replace(ZRENDER_ID_PREFIX, "")}`;

        ids.set(oldId, newId);
        setAttribute(element, "id", newId);
    }

    for (const element of elements) {
        for (const attribute of element.attrs) {
            for (const [oldId, newId] of ids) {
                attribute.value = attribute.value.replaceAll(`url(#${oldId})`, `url(#${newId})`);

                if (attribute.value === `#${oldId}`) {
                    attribute.value = `#${newId}`;
                }
            }
        }
    }

    setAttribute(svg, "width", "100%");
    setAttribute(svg, "height", "100%");

    return element("div", [["class", "slide-element slide-chart"], ...objectAttributes, styleAttribute(css), ["role", "img"], ["aria-label", altText || undefined]], serializeOuter(svg));
}

/** The chart box; addChart() calls it too, so a zero-sized chart fails there rather than at toHtml(). */
export function chartSize(style: NormalizedObjectStyle): { width: number; height: number } {
    const width = style.geometry.width ?? 0;
    const height = style.geometry.height ?? 0;

    if (width <= 0 || height <= 0) {
        throw new RangeError("Chart width and height must be greater than zero");
    }

    return { width, height };
}

export function renderChart(chart: NormalizedChart, style: NormalizedObjectStyle, idScope: string): Html {
    const { width, height } = chartSize(style);

    if (chart.type === "bar3D") {
        return chartContainer(renderBar3DSvg(width, height, chart), style, chart.altText, idScope);
    }

    // The ECharts option of a 2D chart.
    const keySize = Math.round(pointsToPixels(chart.legendFontSize) * PPTX_DEFAULTS.chart.legendKeyFontRatio);
    const hasLineKey = chart.type === "radar" || chart.series.some(series => series.type === "line");
    // LibreOffice's title and top or bottom legend bands (PPTX_DEFAULTS.chart.plotBand): a legend's row is legendBandEm
    // legend font sizes tall with its entries centered in it, a top one legendBandPt below the title band and a bottom
    // one legendBottomPt above the chart's bottom edge. ECharts (padding 0) centers an entry half a font size from the
    // legend's edge, so the edge goes the rest of half a row further in.
    const gap = height * PPTX_DEFAULTS.chart.plotAreaGapRatio;
    const plotBand = PPTX_DEFAULTS.chart.plotBand;
    const pieLayout = PPTX_DEFAULTS.chart.pieLayout;
    const legendPx = pointsToPixels(chart.legendFontSize);
    const legendBand = legendPx * plotBand.legendBandEm + pointsToPixels(plotBand.legendBandPt);
    const legendRowInset = (legendPx * (plotBand.legendBandEm - 1)) / 2;
    // A turned title (titleRotate) takes its turned box's height, as LibreOffice 26.2 lowers the plot under it.
    const titleTurn = ((chart.titleRotate ?? 0) * Math.PI) / 180;
    const titleTextPx = pointsToPixels(chart.titleFontSize);
    const titleTextWidth = chart.showTitle && chart.titleRotate ? format.getTextRect(chart.title, `${titleTextPx}px ${chart.titleFontFace}`).width : 0;
    const titleTurnedWidth = Math.abs(Math.cos(titleTurn)) * titleTextWidth + Math.abs(Math.sin(titleTurn)) * titleTextPx;
    const titleTurnedHeight = chart.titleRotate ? Math.abs(Math.sin(titleTurn)) * titleTextWidth + Math.abs(Math.cos(titleTurn)) * titleTextPx : titleTextPx;
    const titleBand = chart.showTitle ? gap + titleTurnedHeight - titleTextPx + titleTextPx * pieLayout.titleBandEm + pointsToPixels(plotBand.titleBandPt) : 0;
    // A vertical legend's entries come every pitch: ECharts puts them a text line and itemGap apart, and centers
    // the first half a text line below the legend's top (padding 0).
    const verticalLegendLayout = PPTX_DEFAULTS.chart.verticalLegend;
    const topRight = chart.topRightLegend;
    const legendPitch = (topRight.kind === "above" && chart.legendPosition === "tr" ? topRight.pitchEm : verticalLegendLayout.pitchEm) * legendPx;
    const legendItemGap = legendPitch - legendPx;
    let legendPlacement: Record<string, unknown>;

    switch (chart.legendPosition) {
        case "b":
            legendPlacement = { left: "center", bottom: pointsToPixels(plotBand.legendBottomPt) + legendRowInset, orient: "horizontal", padding: 0 };
            break;
        case "l":
            legendPlacement = { left: 0, top: "middle", orient: "vertical", itemGap: legendItemGap };
            break;
        case "t":
            legendPlacement = { left: "center", top: titleBand + pointsToPixels(plotBand.legendBandPt) + legendRowInset, orient: "horizontal", padding: 0 };
            break;
        case "tr":
            // A band above the plot area (below a title's), at the chart's right edge; or the chart's top corner, beside it.
            // shortcut: neither is measured under a title; measure it when a deck has one.
            legendPlacement =
                topRight.kind === "above"
                    ? { right: gap, top: titleBand + (topRight.firstKeyCenterEm - 0.5) * legendPx, orient: "vertical", padding: 0, itemGap: legendItemGap }
                    : { right: 0, top: (verticalLegendLayout.firstKeyCenterEm - 0.5) * legendPx, orient: "vertical", padding: 0, itemGap: legendItemGap };
            break;
        case "r":
            legendPlacement = { right: 0, top: "middle", orient: "vertical", itemGap: legendItemGap };
            break;
    }

    const horizontal = chart.type === "bar" && chart.barDirection === "bar";
    const stacked = chart.grouping === "stacked" || chart.grouping === "percentStacked";
    // A clustered horizontal bar chart runs its series bottom-up in each category band, as its categories run, and lists
    // them in the legend from the last; a stacked one keeps the series order in both (LibreOffice and PowerPoint, #472).
    const reversedSeries = horizontal && !stacked;
    // A vertical legend lists a stacked column or area chart from the last series, as its stack runs bottom-up; a
    // horizontal legend keeps the series order (LibreOffice 26.2, #508).
    const verticalLegend = chart.legendPosition === "l" || chart.legendPosition === "r" || chart.legendPosition === "tr";
    const reversedLegend = reversedSeries || (verticalLegend && stacked && (chart.type === "area" || (chart.type === "bar" && !horizontal)));
    // Office keys a bar, area or pie entry with a filled square and a line entry with its line and marker.
    let legendData: { name: string; icon?: string; itemStyle?: { color: string } }[] | undefined;

    if (chart.type === "pie" || chart.type === "doughnut") {
        const labels = chart.series[0]?.labels ?? [];

        legendData = (chart.series[0]?.values ?? []).map((_, index) => ({ name: labels[index] ?? String(index + 1), icon: LEGEND_SQUARE_ICON }));
    } else if (chart.legendListsPoints) {
        // Each point of the only series, keyed in its drawn fill, as PowerPoint's PDF export draws it (#546, #552);
        // ECharts keys a bar series' legend by series name, so each entry gets an empty series of its name below.
        const only = chart.series[0];

        legendData = [];

        for (const [index] of (only?.values ?? []).entries()) {
            legendData.push({ name: only?.labels[index] ?? String(index + 1), icon: LEGEND_SQUARE_ICON, itemStyle: { color: only?.automaticPointFill ?? only?.pointFills?.[index] ?? chart.colors[index % chart.colors.length] ?? "#4472C4" } });
        }
    } else if (chart.type !== "radar" && chart.type !== "scatter" && chart.type !== "bubble") {
        legendData = [];

        // A combo chart's area series first where the target lists them so (legendListsAreasFirst, #538).
        for (const areas of chart.legendListsAreasFirst ? [true, false] : [undefined]) {
            for (const series of chart.series) {
                if (areas === undefined || areas === (series.type === "area")) {
                    legendData.push(series.type === "line" ? { name: series.name } : { name: series.name, icon: LEGEND_SQUARE_ICON });
                }
            }
        }

        if (reversedLegend) {
            legendData.reverse();
        }
    }

    const legend = {
        show: chart.showLegend,
        ...legendPlacement,
        data: legendData,
        itemHeight: keySize,
        // A line key keeps ECharts' item width; without one every key is a square.
        ...(hasLineKey ? {} : { itemWidth: keySize }),
        textStyle: {
            color: chart.legendColor,
            fontFamily: chart.legendFontFace,
            fontSize: pointsToPixels(chart.legendFontSize),
        },
    };
    // PptxGenJS writes <c:marker> on line, radar and scatter series: <c:size> is in points (ECharts takes pixels) and
    // <a:ln> outlines each marker with lineDataSymbolLineSize in lineDataSymbolLineColor or the series color.
    const symbolPixels = pointsToPixels(chart.lineSymbolSize);
    const symbolSize = chart.lineSymbol === "dash" ? [symbolPixels, symbolPixels * DASH_SYMBOL_HEIGHT_RATIO] : symbolPixels;
    const symbolBorder = { borderWidth: pointsToPixels(chart.lineSymbolLineSize), ...(chart.lineSymbolLineColor ? { borderColor: chart.lineSymbolLineColor } : {}) };
    // The chart's graphic elements: a placed or turned title, a pie's rim labels, a data table.
    const graphics: Record<string, unknown>[] = [];
    // A manually placed title (titlePos) has its box's top left corner there; a turned one (titleRotate, clockwise) turns
    // about its middle, its turned box kept there. ECharts' title component does neither, so the title is graphic text.
    const titleGraphic = chart.showTitle && (chart.titlePosition !== undefined || chart.titleRotate !== undefined);

    if (titleGraphic) {
        const titlePx = pointsToPixels(chart.titleFontSize);
        const position = chart.titlePosition;
        const turned = chart.titleRotate !== undefined && chart.titleRotate !== 0;
        const anchored = position !== undefined && !turned;

        graphics.push({
            type: "text",
            x: position ? position.x * width + (turned ? titleTurnedWidth / 2 : 0) : width / 2,
            y: position ? position.y * height + (turned ? titleTurnedHeight / 2 : 0) : TITLE_PADDING + titleTurnedHeight / 2,
            rotation: (-(chart.titleRotate ?? 0) * Math.PI) / 180,
            silent: true,
            style: {
                text: chart.title,
                fill: chart.titleColor,
                fontFamily: chart.titleFontFace,
                fontSize: titlePx,
                fontWeight: chart.titleBold ? "bold" : "normal",
                align: anchored ? "left" : "center",
                verticalAlign: anchored ? "top" : "middle",
            },
        });
    }

    const common: EChartsCoreOption = {
        animation: false,
        color: chart.colors,
        textStyle: {
            fontFamily: chart.fontFace,
            fontSize: pointsToPixels(chart.fontSize),
        },
        title:
            chart.showTitle && !titleGraphic
                ? {
                      show: true,
                      text: chart.title,
                      left: "center",
                      textStyle: {
                          color: chart.titleColor,
                          fontFamily: chart.titleFontFace,
                          fontSize: pointsToPixels(chart.titleFontSize),
                          fontWeight: chart.titleBold ? "bold" : "normal",
                      },
                  }
                : { show: false },
        legend,
    };
    const axisStyle = {
        axisLabel: {
            color: "#000000",
            fontFamily: chart.fontFace,
            fontSize: pointsToPixels(chart.fontSize),
        },
        // The category axis line (a scatter chart's x axis line): catAxisLine*, none on a deleted axis.
        axisLine: {
            show: chart.axisLineVisible && !chart.categoryAxis.hidden,
            lineStyle: { color: chart.categoryAxis.lineColor, width: pointsToPixels(chart.categoryAxis.lineWidth), ...(chart.categoryAxis.lineStyle === "solid" ? {} : { type: lineType(chart.categoryAxis.lineStyle) }) },
        },
    };
    // The automatic plot area (PptxGenJS writes `<c:layout/>`) in text-sized bands, not fractions of the chart height,
    // as LibreOffice lays it out (PPTX_DEFAULTS.chart.plotBand, and the title and legend bands above).
    const valueLabelPx = pointsToPixels(chart.valueAxes[0]?.labelFontSize ?? chart.fontSize);
    const categoryAxisModel = chart.categoryAxis;
    const categoryLabelPx = pointsToPixels(categoryAxisModel.labelFontSize);
    const bottomLabelPx = horizontal ? valueLabelPx : categoryLabelPx;
    const bottomTickMark = horizontal ? chart.valueAxes[0]?.majorTickMark : chart.categoryMajorTickMark;
    // The bottom axis's labels sit below its outside ticks and the label spacing, centered in a line of labelLineEm
    // label font sizes; ECharts measures axisLabel.margin from the axis line to the label's font box.
    const bottomLabelMargin = (bottomTickMark === "outside" || bottomTickMark === "cross" ? pointsToPixels(plotBand.tickLengthPt) : 0) + pointsToPixels(plotBand.labelSpacingPt) + (bottomLabelPx * (plotBand.labelLineEm - 1)) / 2;
    const lineEm = PPTX_DEFAULTS.text.singleLineEm;
    const inset = pointsToPixels(pieLayout.insetPt);
    // A left or right legend's band beside the plot area, as LibreOffice reserves it on every chart type (#468, #517):
    // its widest entry (key, gap and text) and an inset on each side.
    // shortcut: zrender measures the entries with its Arial-like width table, so a narrower legend face (LibreOffice
    // draws the theme's Calibri as Carlito) leaves the plot a few px smaller; measure the face's advances when charts
    // get them.
    let sideLegendBand = 0;

    if (chart.showLegend && verticalLegend) {
        // A legend without data of its own lists the series names (a scatter chart's y series).
        const entries = legendData ?? (chart.type === "scatter" || chart.type === "bubble" ? chart.series.slice(1) : chart.series);
        let widest = 0;

        for (const entry of entries) {
            widest = Math.max(widest, format.getTextRect(entry.name, `${pointsToPixels(chart.legendFontSize)}px ${chart.legendFontFace}`).width);
        }

        sideLegendBand = 2 * inset + (hasLineKey ? LEGEND_LINE_KEY_WIDTH : keySize) + LEGEND_TEXT_GAP + widest;
    }

    // A left legend moves the value axis gutter and the label bound by its band less one inset: LibreOffice keeps the
    // axis labels an inset clear of the legend.
    const leftLegendShift = chart.legendPosition === "l" ? Math.max(0, sideLegendBand - inset) : 0;
    const topRightAbove = chart.showLegend && chart.legendPosition === "tr" && topRight.kind === "above";
    const rightLegendBand = chart.legendPosition === "r" || (chart.legendPosition === "tr" && !topRightAbove) ? sideLegendBand : 0;
    // A top-right legend above the plot area moves its top from the band's top gap to the band's end below the last key.
    const topLegendBand = topRight.kind === "above" && topRightAbove ? (topRight.firstKeyCenterEm + topRight.bandBelowEm) * legendPx + (Math.max(1, legendData?.length ?? chart.series.length) - 1) * legendPitch - gap : 0;
    // A bottom category-axis title gets its band below the labels, as LibreOffice lays it out; a left one (a horizontal
    // bar chart's) is moved clear of the labels by ECharts (nameMoveOverlap) and kept in the frame by the outer bounds.
    const categoryTitleBand = categoryAxisModel.showTitle && !horizontal ? 2 * gap + pointsToPixels(categoryAxisModel.titleFontSize) * PPTX_DEFAULTS.chart.categoryAxis.titleBandEm : 0;
    // A secondary category axis (#603): its labels in a row of their own under the primary ones, or over them (overlay);
    // `high` puts them above the plot. Its title takes a band beside its row, or above the plot when it overlays.
    const secondaryCategory = chart.secondaryCategoryAxis;
    const secondaryLabelPx = secondaryCategory ? pointsToPixels(secondaryCategory.labelFontSize) : 0;
    const secondaryLabels = secondaryCategory !== undefined && !secondaryCategory.hidden && secondaryCategory.labelPosition !== "none";
    const secondaryHigh = secondaryCategory?.labelPosition === "high";
    const secondaryRow = secondaryCategory?.placement === "row";
    const secondaryLabelBand = secondaryLabels && (secondaryRow || secondaryHigh) ? bottomLabelMargin + secondaryLabelPx * lineEm : 0;
    const secondaryTitleBand = secondaryCategory?.showTitle && !secondaryCategory.hidden ? 2 * gap + pointsToPixels(secondaryCategory.titleFontSize) * PPTX_DEFAULTS.chart.categoryAxis.titleBandEm : 0;
    const secondaryTitleTop = !secondaryRow || secondaryHigh;
    const secondaryBottomBand = (secondaryHigh ? 0 : secondaryLabelBand) + (secondaryTitleTop ? 0 : secondaryTitleBand);
    const secondaryTopBand = (secondaryHigh ? secondaryLabelBand : 0) + (secondaryTitleTop ? secondaryTitleBand : 0);
    // The value axis gutter: room for its labels and title.
    const firstAxis = chart.valueAxes[0];
    let gutter: number;

    if (chart.valueAxes.length > 1) {
        gutter = width * 0.12;
    } else if (!firstAxis?.showTitle) {
        gutter = width * 0.06;
    } else {
        const labels = [firstAxis.minimum, firstAxis.maximum].map(value => formatChartNumber(value, firstAxis.labelFormatCode));
        const labelWidth = Math.max(...labels.map(label => label.length), 1) * pointsToPixels(firstAxis.labelFontSize) * 0.6;

        gutter = Math.ceil(Math.max(80, labelWidth + pointsToPixels(firstAxis.titleFontSize) + 30));
    }

    // The gutters are the least room beside the plot; ECharts shrinks the plot further until every axis label and title
    // measures (zrender's text width, as it lays the labels out) inside the chart less the gap on each side, as
    // PowerPoint and LibreOffice size the plot area so the labels fit the frame (#443): a currency value axis, long
    // category names, the last value label centered on a horizontal bar chart's right edge. Vertically the bands above
    // already hold the labels, so the bounds there are the chart's edges.
    // shortcut: without a canvas (Node, Bun) zrender measures every face with its built-in Arial-like width table, so a
    // face much wider than Arial can still overflow by a few px; measure with the face's advances when charts get them.
    const labelBound = gap - AXIS_LABEL_TEXT_MARGIN;
    const plotArea = chart.plotArea;
    // The plot area's fill and border cover the inner plot rectangle (the grid), as LibreOffice draws c:plotArea/c:spPr.
    const plotAreaStyle = {
        show: plotArea.fill.visible || plotArea.borderVisible,
        backgroundColor: plotArea.fill.visible ? plotArea.fill.color : "transparent",
        borderColor: plotArea.borderColor,
        borderWidth: plotArea.borderVisible ? pointsToPixels(plotArea.borderWidth) : 0,
    };
    // A pie, doughnut or radar chart has no grid to carry the plot area's fill and border, so they are a square behind the
    // ring or the spokes (normalizeChart keeps them only where the target draws them, drawsPlotAreaOnEveryType, #565).
    const plotAreaSquare = (x: number, y: number, side: number) => ({
        type: "rect",
        silent: true,
        shape: { x, y, width: side, height: side },
        style: { fill: plotAreaStyle.backgroundColor, stroke: plotArea.borderVisible ? plotAreaStyle.borderColor : "transparent", lineWidth: plotAreaStyle.borderWidth },
    });
    const layout = chart.plotLayout;
    // A data table (#567) takes the category labels' band: a row of category names, then one per series, each
    // rowEm table font sizes tall, its series names in a column left of the plot area. Its columns are the categories',
    // so the plot area is laid out here (outerBoundsMode none), its gutter wide enough for the names and value labels.
    const table = chart.dataTable;
    const tableDefaults = PPTX_DEFAULTS.chart.dataTable;
    const tablePx = table ? pointsToPixels(table.fontSize) : 0;
    const tableRowPx = tablePx * tableDefaults.rowEm;
    const tableKeyPx = tablePx * tableDefaults.keyEm;
    let tableNameColumn = 0;

    if (table) {
        let widestName = 0;
        let widestValueLabel = 0;

        for (const series of chart.series) {
            widestName = Math.max(widestName, format.getTextRect(series.name, `${tablePx}px ${table.fontFace}`).width);
        }

        for (const value of firstAxis ? [firstAxis.minimum, firstAxis.maximum] : []) {
            widestValueLabel = Math.max(widestValueLabel, format.getTextRect(formatChartNumber(value, firstAxis?.labelFormatCode ?? "General"), `${valueLabelPx}px ${firstAxis?.labelFontFace ?? chart.fontFace}`).width);
        }

        // LibreOffice 26.2 starts the names two text gaps in and keeps the plot area one gap after the widest.
        tableNameColumn = 3 * LEGEND_TEXT_GAP + (table.keys ? tableKeyPx + LEGEND_TEXT_GAP : 0) + widestName;
        gutter = Math.max(gutter, tableNameColumn, widestValueLabel + TICK_LENGTH + AXIS_LABEL_MARGIN + gap);
    }

    // A column chart whose category labels are not below it (deleted, `none` or `high`) gives their band to the plot, as
    // LibreOffice 26.2 does.
    const noBottomLabels = !horizontal && chart.type !== "scatter" && chart.type !== "bubble" && (categoryAxisModel.hidden || categoryAxisModel.labelPosition === "none" || (categoryAxisModel.labelPosition === "high" && !firstAxis?.reversed));
    const bottomLabelBand = bottomLabelMargin + (bottomLabelPx * (plotBand.labelLineEm + 1)) / 2;
    // A horizontal bar chart's data table goes below its value labels, as LibreOffice 26.2.6 draws it (#583).
    const tableOffset = table && horizontal ? bottomLabelBand : 0;
    const categoryBand = table ? (chart.series.length + 1) * tableRowPx + tableOffset : noBottomLabels ? 0 : bottomLabelBand;
    // A column chart's value axis on the right (catAxisCrossesAt max, or a reversed category axis) takes the gutter there.
    const valueAxisRight =
        !horizontal && !table && chart.valueAxes.length === 1 && chart.type !== "scatter" && chart.type !== "bubble" && chart.type !== "radar" && (categoryAxisModel.valueAxisCrossing?.kind === "max") !== (categoryAxisModel.reversed === true);
    const gridLeft = valueAxisRight ? width * 0.02 + leftLegendShift : gutter + leftLegendShift;
    // A right legend's band replaces the 2% margin; a secondary axis keeps its gutter beside the band.
    // shortcut: the secondary axis beside a right legend is not measured against LibreOffice; measure it when a deck
    // needs both.
    const gridRight = chart.valueAxes.length > 1 ? width * 0.14 + rightLegendBand : valueAxisRight ? gutter + rightLegendBand : Math.max(width * 0.02, rightLegendBand);
    const gridBottom = gap + (chart.showLegend && chart.legendPosition === "b" ? legendBand : 0) + categoryTitleBand + categoryBand + secondaryBottomBand;
    const gridTop = gap + titleBand + (chart.showLegend && chart.legendPosition === "t" ? legendBand : 0) + (horizontal ? 0 : valueLabelPx / 2) + topLegendBand + secondaryTopBand;
    // A manual layout (layoutTarget inner) places the plot rectangle itself, so the labels go where they fall.
    const grid = layout
        ? { left: layout.x * width, top: layout.y * height, width: layout.w * width, height: layout.h * height, outerBoundsMode: "none", ...plotAreaStyle }
        : {
              left: gridLeft,
              right: gridRight,
              top: gridTop,
              bottom: gridBottom,
              outerBoundsMode: table ? "none" : "auto",
              outerBounds: { left: labelBound + leftLegendShift, right: Math.max(labelBound, rightLegendBand), top: 0, bottom: 0 },
              ...plotAreaStyle,
          };
    // An automatic value axis (no valAxisMajorUnit) takes the steps its target fits along the drawn axis (stepRoom, #585):
    // at most the plot's length over the room per step, in label heights or the first scale's widest label widths.
    // shortcut: ECharts may still narrow the plot for its labels (outerBounds), so a horizontal axis's length can be
    // long: 18pt labels on 0.3..0.9 step by 0.2 on an axis LibreOffice draws 110 to 120pt long, where it steps by 0.5.
    // Take the laid-out length (a second layout pass, as LibreOffice's) when a deck needs it (#585).
    const plotWidth = layout ? layout.w * width : width - gridLeft - gridRight;
    const valueAxisLength = horizontal ? plotWidth : layout ? layout.h * height : height - gridTop - gridBottom;
    const scaledValueAxes: NormalizedChartValueAxis[] = [];
    // How many steps a room fits along an axis of that length: the labels' height, or the widest label of the first
    // scale (minimum to maximum by its automatic interval) in the axis's format, per step.
    const stepRoomIncrements = (
        room: NonNullable<NonNullable<NormalizedChartValueAxis["automaticScale"]>["stepRoom"]>,
        length: number,
        axis: { minimum: number; maximum: number; labelFontSize: number; labelFontFace: string; labelFormatCode: string; divisor: number },
    ) => {
        const labelPx = pointsToPixels(axis.labelFontSize);
        let labelSize = 0;

        if (room.label === "height") {
            const font = builtInFontMetrics(axis.labelFontFace);

            labelSize = labelPx * (font.kind === "metrics" ? (font.metrics.winAscent + font.metrics.winDescent) / font.metrics.unitsPerEm : PPTX_DEFAULTS.text.singleLineEm);
        } else {
            const interval = automaticAxisInterval(axis.minimum, axis.maximum);

            for (let step = 0; axis.minimum + step * interval <= axis.maximum + interval / 2; step++) {
                const value = Number((axis.minimum + step * interval).toPrecision(12));

                labelSize = Math.max(labelSize, format.getTextRect(formatChartNumber(value / axis.divisor, axis.labelFormatCode), `${labelPx}px ${axis.labelFontFace}`).width);
            }
        }

        return length / (room.perStep * Math.max(1, labelSize));
    };
    // LibreOffice's rule takes whole steps, between 2 and the room's maxIncrements.
    const wholeSteps = (room: { rounding: number; maxIncrements: number }, steps: number) => Math.max(2, Math.min(room.maxIncrements, Math.floor(steps + room.rounding)));

    for (const [index, axis] of chart.valueAxes.entries()) {
        const automatic = axis.automaticScale;
        const room = automatic?.stepRoom;
        // A percent-stacked axis runs 0 to 100 whatever its data.
        const percent = chart.series.some(series => series.stack === "percentStacked" && series.valueAxisIndex === index);

        if (!automatic || !room || percent) {
            scaledValueAxes.push(axis);
            continue;
        }

        const roomSteps = stepRoomIncrements(room, valueAxisLength, { ...axis, divisor: axis.displayUnit?.divisor ?? 1 });
        const fixed = automatic.fixedMinimum ? { kind: "minimum" as const, value: axis.minimum } : automatic.fixedMaximum ? { kind: "maximum" as const, value: axis.maximum } : { kind: "none" as const };
        const scale = automatic.padding
            ? paddedAxisScale(automatic.low, automatic.high, fixed, automatic.padding, { kind: "room", steps: roomSteps })
            : automaticAxisScale(automatic.low, automatic.high, fixed, true, { kind: "automatic", maxIncrements: wholeSteps(room, roomSteps) });

        scaledValueAxes.push({ ...axis, minimum: automatic.fixedMinimum ? axis.minimum : scale.minimum, maximum: automatic.fixedMaximum ? axis.maximum : scale.maximum, majorUnit: scale.interval });
    }

    // The category axis's labels and title (a scatter or bubble chart's x axis takes them too, as PptxGenJS writes its
    // c:valAx from the catAxis* options).
    const categoryLabelFormat = categoryAxisModel.labelFormatCode;
    const categoryAxisText = {
        axisLabel: {
            // A deleted axis, `none` and a column chart's data table (whose first row names the categories) print no
            // labels; a horizontal bar chart keeps them left of the plot, as LibreOffice 26.2.6 does (#583).
            ...(categoryAxisModel.hidden || categoryAxisModel.labelPosition === "none" || (chart.dataTable && !horizontal) ? { show: false } : {}),
            // catAxisLabelRotate turns clockwise, ECharts' rotate counterclockwise.
            ...(categoryAxisModel.labelRotate ? { rotate: -categoryAxisModel.labelRotate } : {}),
            ...(categoryAxisModel.labelSkip ? { interval: categoryAxisModel.labelSkip - 1 } : {}),
            // catLabelFormatCode formats each numeric category (PptxGenJS writes them as a date axis's numbers).
            ...(categoryLabelFormat
                ? {
                      formatter: (value: string) => {
                          const number = Number(value);

                          return value.trim() !== "" && Number.isFinite(number) ? formatChartNumber(number, categoryLabelFormat) : value;
                      },
                  }
                : {}),
            color: categoryAxisModel.labelColor,
            fontFamily: categoryAxisModel.labelFontFace,
            fontSize: categoryLabelPx,
            // shortcut: a scatter x axis keeps ECharts' label margin, as its spacing is not measured against LibreOffice.
            ...(horizontal || chart.type === "scatter" || chart.type === "bubble" ? {} : { margin: bottomLabelMargin }),
            ...(categoryAxisModel.labelBold ? { fontWeight: "bold" } : {}),
            ...(categoryAxisModel.labelItalic ? { fontStyle: "italic" } : {}),
        },
        // LibreOffice 26.2 draws a category axis's minor ticks at the category boundaries, as its major ones, so either
        // draws them there.
        axisTick: categoryAxisModel.hidden ? { show: false } : tickOption(chart.categoryMajorTickMark !== "none" ? chart.categoryMajorTickMark : categoryAxisModel.minorTickMark, categoryAxisModel.lineColor, categoryAxisModel.lineWidth),
        name: categoryAxisModel.showTitle && !categoryAxisModel.hidden ? categoryAxisModel.title : undefined,
        nameLocation: "middle",
        nameGap: horizontal ? AXIS_LABEL_MARGIN : TICK_LENGTH + AXIS_LABEL_MARGIN + categoryLabelPx * lineEm + gap + (secondaryHigh ? 0 : secondaryLabelBand),
        nameRotate: categoryAxisModel.titleRotate,
        nameTextStyle: {
            color: categoryAxisModel.titleColor,
            fontFamily: categoryAxisModel.titleFontFace,
            fontSize: pointsToPixels(categoryAxisModel.titleFontSize),
        },
    } as const;
    let option: EChartsCoreOption;

    if (chart.type === "pie" || chart.type === "doughnut") {
        const [series] = chart.series;

        if (!series) {
            throw new TypeError(`A ${chart.type} chart model needs one data series; got none`);
        }

        // Each slice's label: its series name, category, value and share of the total, each shown field on its own line.
        const labels = chart.dataLabels;
        let total = 0;

        for (const value of series.values) {
            total += value;
        }

        const sliceLabel = (index: number) => {
            const value = series.values[index] ?? 0;
            const fields: string[] = [];

            if (labels.showSeriesName) {
                fields.push(series.name);
            }

            if (labels.showCategory) {
                fields.push(series.labels[index] ?? String(index + 1));
            }

            if (labels.showValue) {
                fields.push(formatChartNumber(value, labels.valueFormat));
            }

            if (labels.showPercent) {
                fields.push(formatChartNumber(total === 0 ? 0 : value / total, labels.percentFormat));
            }

            return fields.join(PPTX_DEFAULTS.chart.dataLabelSeparator);
        };

        // The ring's box: the plot area the title and legend leave, less the inset (PPTX_DEFAULTS.chart.pieLayout).
        let top: number = chart.showTitle ? gap + pointsToPixels(chart.titleFontSize) * pieLayout.titleBandEm + pointsToPixels(pieLayout.titleBandPt) : inset;
        let bottom: number = inset;
        let left: number = inset;
        let right: number = inset;

        if (chart.showLegend) {
            if (chart.legendPosition === "b") {
                bottom = 2 * inset + legendPx * pieLayout.legendBandEm;
            } else if (chart.legendPosition === "t") {
                top += inset + legendPx * pieLayout.legendBandEm;
            } else if (chart.legendPosition === "l") {
                left = sideLegendBand;
            } else {
                right = sideLegendBand;
            }
        }

        // A manual layout (layoutTarget inner) is the ring's box itself.
        if (layout) {
            left = layout.x * width;
            top = layout.y * height;
            right = width - left - layout.w * width;
            bottom = height - top - layout.h * height;
        }

        const boxWidth = width - left - right;
        const boxHeight = height - top - bottom;
        let radius = Math.max(0, Math.min(boxWidth, boxHeight) / 2);
        let centerX = left + boxWidth / 2;
        let centerY = top + boxHeight / 2;
        const showLabels = labels.showSeriesName === true || labels.showCategory || labels.showValue || labels.showPercent;
        // inEnd and outEnd labels sit at the rim of each slice, and bestFit ones inside it near the rim
        // (PPTX_DEFAULTS.chart.pieLabel); ECharts' pie label positions have none of them, so they are drawn as text.
        const rimLabels = showLabels && (labels.position === "inEnd" || labels.position === "outEnd");
        const bestFit = showLabels && labels.position === "bestFit";
        const rimTexts: Record<string, unknown>[] = [];

        if (rimLabels || bestFit) {
            const pieLabel = PPTX_DEFAULTS.chart.pieLabel;
            // A bestFit label that does not fit in its slice goes outside it, as an outEnd one.
            const outside = labels.position !== "inEnd";
            const offset = pointsToPixels(pieLabel.offsetPt) * (outside ? 1 : -1);
            const labelPx = pointsToPixels(labels.fontSize);
            const lineHeight = labelPx * pieLabel.lineEm;
            const font = `${labels.italic ? "italic " : ""}${labels.bold ? "bold " : ""}${labelPx}px ${labels.fontFace}`;
            const band = pieLabel.axisBandDeg;
            const wrapWidth = chart.outsidePieLabelWidth;
            // A label's fields, each on its own lines: an outside label wraps where the target wraps it
            // (outsidePieLabelWidth) beside its slice's rim point at rimX, between words, or between characters in a word
            // too long for its line; rimX undefined keeps one line per field. PieChart::createTextLabelShape takes the room
            // on the right as the plot area's width less rimX, both from the chart's left edge, and on the left as rimX.
            const layoutLabel = (fields: readonly string[], cos: number, rimX: number | undefined) => {
                const maxWidth = wrapWidth && rimX !== undefined ? Math.min(wrapWidth.room * (cos > 0 ? boxWidth - rimX : rimX), wrapWidth.plotWidth * boxWidth) : Infinity;
                const lines: string[] = [];
                let width = 0;

                for (const field of fields) {
                    let line = "";

                    for (const char of field) {
                        if (line && format.getTextRect(`${line}${char}`, font).width > maxWidth) {
                            const space = line.lastIndexOf(" ");

                            lines.push(space > 0 ? line.slice(0, space) : line);
                            line = space > 0 ? `${line.slice(space + 1)}${char}` : char;
                        } else {
                            line += char;
                        }
                    }

                    lines.push(line);
                }

                for (const line of lines) {
                    width = Math.max(width, format.getTextRect(line, font).width);
                }

                return { text: lines.join("\n"), width, height: lines.length * lineHeight };
            };
            // Each slice's label fields, bisector (degrees counterclockwise from 3 o'clock; ECharts lays the slices out
            // clockwise from startAngle), angular width and its side of the anchor.
            const slices: { fields: string[]; degrees: number; spanDegrees: number; cos: number; sin: number; align: "left" | "center" | "right"; verticalAlign: "top" | "middle" | "bottom" }[] = [];
            let before = 0;

            for (const [index, value] of series.values.entries()) {
                const share = total === 0 ? 1 / series.values.length : value / total;
                const degrees = (((90 - chart.firstSliceAngle - (before + share / 2) * 360) % 360) + 360) % 360;
                // LibreOffice puts an outEnd label right of its anchor on the right half and above it on the upper
                // half, an inEnd label on the other side, and centers it on the bisector near an axis.
                const rightHalf = degrees < 90 - band || degrees > 270 + band;
                const leftHalf = degrees > 90 + band && degrees < 270 - band;
                const upperHalf = degrees > band && degrees < 180 - band;
                const lowerHalf = degrees > 180 + band && degrees < 360 - band;
                const radians = (degrees * Math.PI) / 180;

                before += share;
                slices.push({
                    fields: sliceLabel(index).split(PPTX_DEFAULTS.chart.dataLabelSeparator),
                    degrees,
                    spanDegrees: share * 360,
                    cos: Math.cos(radians),
                    sin: Math.sin(radians),
                    align: (outside ? rightHalf : leftHalf) ? "left" : (outside ? leftHalf : rightHalf) ? "right" : "center",
                    verticalAlign: (outside ? upperHalf : lowerHalf) ? "bottom" : (outside ? lowerHalf : upperHalf) ? "top" : "middle",
                });
            }

            // PieChart::performLabelBestFitInnerPlacement on a ring of that radius: how far from the center the label's
            // center goes, with its outer corner on the ring shrunk by bestFitRimRatio, or undefined when a corner then
            // leaves the slice.
            const bestFitDistance = (slice: (typeof slices)[number], inner: { width: number; height: number }, ringRadius: number) => {
                const halfWidth = inner.width / 2;
                const halfHeight = inner.height / 2;
                const reach = ringRadius * pieLabel.bestFitRimRatio;
                // The outer corner's offset from the label's center, along the bisector.
                const along = Math.abs(slice.cos) * halfWidth + Math.abs(slice.sin) * halfHeight;
                const discriminant = along * along - halfWidth * halfWidth - halfHeight * halfHeight + reach * reach;

                if (discriminant < 0 || Math.sqrt(discriminant) < along) {
                    return undefined;
                }

                const distance = Math.sqrt(discriminant) - along;

                for (const [dx, dy] of [
                    [-halfWidth, -halfHeight],
                    [-halfWidth, halfHeight],
                    [halfWidth, -halfHeight],
                    [halfWidth, halfHeight],
                ] as const) {
                    const cornerDegrees = (Math.atan2(distance * slice.sin + dy, distance * slice.cos + dx) * 180) / Math.PI;
                    const fromStart = (((cornerDegrees - slice.degrees + slice.spanDegrees / 2) % 360) + 360) % 360;

                    if (fromStart > slice.spanDegrees) {
                        return undefined;
                    }
                }

                return distance;
            };
            // LibreOffice's first layout: a ring 1 / firstPassDivisor of the plot area, centered in a box that size
            // placed at the plot area's far corner (VDiagram::reduceToMinimumSize), and the rect its labels and ring
            // cover there. inEnd and outEnd labels resize the ring, and on some targets (bestFitLabelsShrinkPie) so does each bestFit label that does not fit in
            // its slice of that small ring, though it may fit in the final one: LibreOffice 26.2 draws a 70/15/8/4/3 pie
            // with bestFit percentages, all inside their slices, on a ring 6% smaller than a 40/25/20/15 one (#572).
            const firstWidth = boxWidth / pieLabel.firstPassDivisor;
            const firstHeight = boxHeight / pieLabel.firstPassDivisor;
            const firstSide = Math.min(firstWidth, firstHeight);
            const firstX = left + firstWidth + (firstWidth - firstSide) / 2;
            const firstY = top + firstHeight + (firstHeight - firstSide) / 2;
            let minX = firstX;
            let maxX = firstX + firstSide;
            let minY = firstY;
            let maxY = firstY + firstSide;

            for (const slice of slices) {
                if (bestFit && bestFitDistance(slice, layoutLabel(slice.fields, slice.cos, undefined), firstSide / 2) !== undefined) {
                    continue;
                }

                const anchorX = firstX + firstSide / 2 + (firstSide / 2 + offset) * slice.cos;
                const anchorY = firstY + firstSide / 2 - (firstSide / 2 + offset) * slice.sin;
                const label = layoutLabel(slice.fields, slice.cos, outside ? firstX + (firstSide / 2) * (1 + slice.cos) : undefined);
                const labelLeft = slice.align === "left" ? anchorX : slice.align === "right" ? anchorX - label.width : anchorX - label.width / 2;
                const labelTop = slice.verticalAlign === "top" ? anchorY : slice.verticalAlign === "bottom" ? anchorY - label.height : anchorY - label.height / 2;

                minX = Math.min(minX, labelLeft);
                maxX = Math.max(maxX, labelLeft + label.width);
                minY = Math.min(minY, labelTop);
                maxY = Math.max(maxY, labelTop + label.height);
            }

            // VDiagram::adjustInnerSize: the ring's box grows by the room left beside that rect, to at least a third of
            // the plot area, and moves to keep the labels in the plot area; the ring is the box's centered square.
            let deltaWidth = boxWidth - (maxX - minX);
            let deltaHeight = boxHeight - (maxY - minY);

            if (firstSide + deltaWidth < boxWidth / 3) {
                deltaWidth = boxWidth / 3 - firstSide;
            }

            if (firstSide + deltaHeight < boxHeight / 3) {
                deltaHeight = boxHeight / 3 - firstSide;
            }

            const diffLeft = minX - left;
            const diffRight = left + boxWidth - maxX;
            const diffUp = minY - top;
            const diffDown = top + boxHeight - maxY;
            let innerX = firstX;
            let innerY = firstY;

            if (diffLeft >= 0) {
                innerX -= diffLeft;
            } else if (diffRight >= 0) {
                innerX += diffRight > -diffLeft ? -diffLeft : diffRight > Math.abs(deltaWidth) ? diffRight : Math.abs(deltaWidth);
            }

            if (diffUp >= 0) {
                innerY -= diffUp;
            } else if (diffDown >= 0) {
                innerY += diffDown > -diffUp ? -diffUp : diffDown > Math.abs(deltaHeight) ? diffDown : Math.abs(deltaHeight);
            }

            // Clipped to the plot area (VDiagram::adjustPosAndSize).
            const innerLeft = Math.max(left, innerX);
            const innerTop = Math.max(top, innerY);
            const innerWidth = Math.max(0, Math.min(left + boxWidth, innerX + firstSide + deltaWidth) - innerLeft);
            const innerHeight = Math.max(0, Math.min(top + boxHeight, innerY + firstSide + deltaHeight) - innerTop);

            if (rimLabels || (bestFit && chart.bestFitLabelsShrinkPie)) {
                radius = Math.min(innerWidth, innerHeight) / 2;
                centerX = innerLeft + innerWidth / 2;
                centerY = innerTop + innerHeight / 2;
            }

            for (const slice of slices) {
                let x = centerX + (radius + offset) * slice.cos;
                let y = centerY - (radius + offset) * slice.sin;
                let label = layoutLabel(slice.fields, slice.cos, outside ? centerX + radius * slice.cos : undefined);
                let align = slice.align;
                let verticalAlign = slice.verticalAlign;

                // PieChart::performLabelBestFitInnerPlacement: centered on the bisector with its outer corner on the ring
                // shrunk by bestFitRimRatio, if every corner then lies in the slice; else outside, as above.
                if (bestFit) {
                    const inner = layoutLabel(slice.fields, slice.cos, undefined);
                    const distance = bestFitDistance(slice, inner, radius);

                    if (distance !== undefined) {
                        x = centerX + distance * slice.cos;
                        y = centerY - distance * slice.sin;
                        label = inner;
                        align = "center";
                        verticalAlign = "middle";
                    } else if (!chart.bestFitLabelsShrinkPie) {
                        // A ring that keeps its size leaves no room outside, so the label moves back into the chart, as
                        // PowerPoint keeps "Amet 3%" of a 70/15/8/4/3 pie inside its frame (#572).
                        const left = align === "left" ? x : align === "right" ? x - label.width : x - label.width / 2;
                        const top = verticalAlign === "top" ? y : verticalAlign === "bottom" ? y - label.height : y - label.height / 2;

                        x += Math.max(0, -left) - Math.max(0, left + label.width - width);
                        y += Math.max(0, -top) - Math.max(0, top + label.height - height);
                    }
                }

                rimTexts.push({
                    type: "text",
                    x,
                    y,
                    silent: true,
                    // Over the slices, which ECharts draws at z 2.
                    z: 3,
                    style: {
                        text: label.text,
                        fill: labels.color,
                        fontFamily: labels.fontFace,
                        fontSize: labelPx,
                        fontWeight: labels.bold ? "bold" : "normal",
                        fontStyle: labels.italic ? "italic" : "normal",
                        lineHeight,
                        align,
                        verticalAlign,
                    },
                });
            }
        }

        if (plotAreaStyle.show) {
            graphics.unshift(plotAreaSquare(centerX - radius, centerY - radius, 2 * radius));
        }

        option = {
            ...common,
            ...(graphics.length || rimTexts.length ? { graphic: [...graphics, ...rimTexts] } : {}),
            series: [
                {
                    name: series.name,
                    type: "pie",
                    center: [centerX, centerY],
                    // OOXML's c:holeSize is a percentage of the outer diameter.
                    radius: chart.type === "doughnut" ? [(radius * chart.doughnutHoleSize) / 100, radius] : radius,
                    startAngle: 90 - chart.firstSliceAngle,
                    itemStyle: { opacity: chart.colorOpacity, ...(chart.dataBorder ? { borderColor: chart.dataBorder.color, borderWidth: pointsToPixels(chart.dataBorder.width) } : {}) },
                    // PptxGenJS 4.0.1 writes a `c:dLbl` for every point with showCatName, showVal and showPercent
                    // off unless showLabel, showValue or showPercent is set, so PowerPoint and LibreOffice draw no
                    // slice labels by default; ECharts would draw category names on leaders. A ctr label sits at
                    // mid-radius, ECharts' inside position.
                    label: showLabels && !rimLabels && !bestFit ? dataLabelOption(labels, "inside", sliceLabel) : { show: false },
                    labelLine: { show: false },
                    data: series.values.map((value, index) => ({
                        value,
                        name: series.labels[index] ?? String(index + 1),
                    })),
                },
            ],
        };
    } else if (chart.type === "radar") {
        const axis = chart.valueAxes[0];

        if (!axis) {
            throw new TypeError("A radar chart model needs a value axis; got none");
        }

        const categoryCount = Math.max(...chart.series.map(series => Math.max(series.labels.length, series.values.length)));
        // Every spoke shares the value axis's scale; LibreOffice numbers the first one only.
        const interval = axis.majorUnit ?? automaticAxisInterval(axis.minimum, axis.maximum);
        const valueLabel = {
            show: !axis.hidden && axis.labelPosition !== "none",
            color: axis.labelColor,
            fontFamily: axis.labelFontFace,
            fontSize: pointsToPixels(axis.labelFontSize),
            fontWeight: axis.labelBold ? "bold" : "normal",
            fontStyle: axis.labelItalic ? "italic" : "normal",
            // valAxisDisplayUnit divides the ring labels where the target reads display units.
            formatter: (value: number) => formatChartNumber(axis.displayUnit ? value / axis.displayUnit.divisor : value, axis.labelFormatCode),
        };
        // LibreOffice 26.2.6 numbers the spoke a numeric catAxisCrossesAt names (the first by default, and for `max`).
        const valueSpoke = categoryAxisModel.valueAxisCrossing?.kind === "category" ? categoryAxisModel.valueAxisCrossing.index : 0;
        const indicators = Array.from({ length: categoryCount }, (_, index) => ({
            name: chart.series[0]?.labels[index] ?? String(index + 1),
            min: axis.minimum,
            max: axis.maximum,
            ...(index === valueSpoke ? { axisLabel: valueLabel } : {}),
        }));
        const labels = chart.dataLabels;
        const label = showsPointLabel(labels)
            ? {
                  ...dataLabelOption(labels, POINT_LABEL_POSITIONS[labels.position] ?? "right", () => ""),
                  // A radar label's params name the series (dataIndex) and the spoke (dimensionIndex).
                  formatter: (params: { dataIndex: number; dimensionIndex: number }) => {
                      const series = chart.series[params.dataIndex];

                      return series ? pointLabelText(labels, series.name, series.values[params.dimensionIndex] ?? 0, { kind: "none" }) : "";
                  },
              }
            : undefined;
        // ECharts draws a radar label on its point's symbol, so a series without markers keeps an invisible one.
        const hiddenSymbol = chart.radarStyle === "standard" && label !== undefined;
        // The radar's circle: the manual layout's, or ECharts' default (centered, half the half frame).
        const radarCenter: [number, number] = layout ? [(layout.x + layout.w / 2) * width, (layout.y + layout.h / 2) * height] : [width / 2, height / 2];
        const radarRadius = layout ? Math.min(layout.w * width, layout.h * height) / 2 : (Math.min(width, height) / 2) * RADAR_RADIUS_RATIO;

        if (plotAreaStyle.show) {
            graphics.unshift(plotAreaSquare(radarCenter[0] - radarRadius, radarCenter[1] - radarRadius, 2 * radarRadius));
        }

        option = {
            ...common,
            ...(graphics.length ? { graphic: graphics } : {}),
            radar: {
                indicator: indicators,
                // Office draws the categories clockwise from the top, counterclockwise for catAxisOrientation maxMin.
                clockwise: categoryAxisModel.reversed !== true,
                splitNumber: Math.max(1, Math.round((axis.maximum - axis.minimum) / interval)),
                center: radarCenter,
                radius: radarRadius,
                axisName: {
                    // A deleted category axis and catAxisLabelPos none name no spoke; low and high name them as nextTo.
                    show: !categoryAxisModel.hidden && categoryAxisModel.labelPosition !== "none",
                    color: categoryAxisModel.labelColor,
                    fontFamily: categoryAxisModel.labelFontFace,
                    fontSize: categoryLabelPx,
                    fontWeight: categoryAxisModel.labelBold ? "bold" : "normal",
                    fontStyle: categoryAxisModel.labelItalic ? "italic" : "normal",
                },
                axisLabel: { show: false },
                // The value gridlines are the rings, the value axis line the spokes.
                splitLine: gridLineOption(axis),
                splitArea: { show: false },
                axisLine: {
                    show: axis.lineVisible,
                    lineStyle: { color: axis.lineColor, width: pointsToPixels(axis.lineWidth), type: lineType(axis.lineStyle) },
                },
            },
            series: [
                {
                    type: "radar",
                    symbol: chart.radarStyle === "standard" && !hiddenSymbol ? "none" : lineSymbol(chart.lineSymbol),
                    symbolSize: hiddenSymbol ? 0 : symbolSize,
                    itemStyle: symbolBorder,
                    lineStyle: { width: pointsToPixels(chart.lineSize) },
                    areaStyle: chart.radarStyle === "filled" ? { opacity: chart.colorOpacity } : undefined,
                    label,
                    data: chart.series.map(series => ({ name: series.name, value: series.values })),
                },
            ],
        };
    } else if (chart.type === "scatter" || chart.type === "bubble") {
        // The first series holds the x values; each other series is a y series.
        const xValues = chart.series[0]?.values ?? [];
        const ySeries = chart.series.slice(1);
        const bubbleSizes: number[] = [];

        for (const series of ySeries) {
            bubbleSizes.push(...(series.sizes ?? []));
        }

        const maxBubble = Math.max(1, ...bubbleSizes);
        // LibreOffice scales the x axis to the x values (not to zero), and the y axis to the value axis's bounds by its
        // major unit or the same automatic interval. A bubble chart's x axis takes the same scale, with no room for the
        // bubbles: its PDF export numbers x values 1..4 0.5..4.5 by 0.5, 10..40 5..45 by 5, 0..100 0..120 by 20, -5..5 -6..6
        // by 2 and 1000..4000 500..4500 by 500, whatever the sizes (#548).
        // catAxisMinVal, catAxisMaxVal and catAxisMajorUnit fix the x axis's bounds and step, the rest automatic.
        const fixedX = categoryAxisModel.xScale;
        const xFixed: FixedAxisBound = fixedX?.minimum !== undefined ? { kind: "minimum", value: fixedX.minimum } : fixedX?.maximum !== undefined ? { kind: "maximum", value: fixedX.maximum } : { kind: "none" };
        // A fixed catAxisMajorUnit is the step the automatic bounds round out to, as on a value axis (#595).
        const xStep = fixedX?.majorUnit !== undefined ? { kind: "fixed" as const, interval: fixedX.majorUnit } : undefined;
        // A target may scale it as a horizontal value axis instead, 0 included, in the steps that fit along it (#616).
        const xLow = Math.min(...xValues);
        const xHigh = Math.max(...xValues);
        const xRoom = fixedX?.stepRoom;
        const xRoomSteps = xRoom
            ? stepRoomIncrements(xRoom, plotWidth, {
                  ...automaticAxisScale(xLow, xHigh, xFixed, true),
                  labelFontSize: categoryAxisModel.labelFontSize,
                  labelFontFace: categoryAxisModel.labelFontFace,
                  labelFormatCode: categoryAxisModel.labelFormatCode ?? "General",
                  divisor: 1,
              })
            : 0;
        const xPadding = fixedX?.padding;
        const automaticX =
            xValues.length === 0
                ? undefined
                : xPadding
                  ? paddedAxisScale(xLow, xHigh, xFixed, xPadding, xStep ?? { kind: "room", steps: xRoomSteps })
                  : automaticAxisScale(xLow, xHigh, xFixed, xRoom !== undefined, xStep ?? (xRoom ? { kind: "automatic", maxIncrements: wholeSteps(xRoom, xRoomSteps) } : undefined));
        const xMinimum = fixedX?.minimum ?? automaticX?.minimum;
        const xMaximum = fixedX?.maximum ?? automaticX?.maximum;
        const xScale =
            xMinimum !== undefined && xMaximum !== undefined
                ? { minimum: xMinimum, maximum: xMaximum, interval: fixedX?.majorUnit ?? (fixedX?.minimum !== undefined && fixedX.maximum !== undefined ? automaticAxisInterval(xMinimum, xMaximum) : (automaticX?.interval ?? 1)) }
                : undefined;
        const scatter: Record<string, unknown>[] = [];
        // Where each axis crosses the other (catAxisCrossesAt on the x axis, valAxisCrossesAt on the y axis, #591): at an
        // end, or a value inside, the zero of a hidden copy of the crossed axis shifted by it (ECharts puts an axis line
        // on another axis's zero only), as on a category chart. LibreOffice 26.2.6 and PowerPoint draw the y axis at x 3
        // and the x axis at y 20 for catAxisCrossesAt 3 and valAxisCrossesAt 20; PowerPoint draws a bubble chart's y axis
        // half a category early (bubbleCategoryCrossingShift, applied in the model). shortcut: the labels stay at the plot's edge, where
        // LibreOffice draws them beside the moved line (#584).
        const yAxisModel = scaledValueAxes[0];
        const yAxisCrossing = categoryAxisModel.valueAxisCrossing;
        const xAxisCrossing = yAxisModel?.categoryAxisCrossing;
        const hiddenXAxes: Record<string, unknown>[] = [];
        const hiddenYAxes: Record<string, unknown>[] = [];
        // The y axis: at the x axis's high end, or on a hidden x axis's zero.
        let yAxisPlacement: { kind: "default" } | { kind: "high" } | { kind: "on-zero" } = { kind: "default" };
        // The x axis: at the y axis's low or high end, or on a hidden y axis's zero.
        let xAxisPlacement: { kind: "default" } | { kind: "low" } | { kind: "high" } | { kind: "on-zero" } = { kind: "default" };

        if (yAxisCrossing?.kind === "max" || (yAxisCrossing?.kind === "value" && xScale && yAxisCrossing.value >= xScale.maximum)) {
            yAxisPlacement = { kind: "high" };
        } else if (yAxisCrossing?.kind === "value" && xScale && yAxisCrossing.value > xScale.minimum) {
            hiddenXAxes.push({ type: "value", show: false, min: xScale.minimum - yAxisCrossing.value, max: xScale.maximum - yAxisCrossing.value, inverse: categoryAxisModel.reversed === true, splitLine: { show: false } });
            yAxisPlacement = { kind: "on-zero" };
        }

        if (yAxisModel && xAxisCrossing !== undefined) {
            const low = yAxisModel.minimum;
            const high = yAxisModel.maximum;
            const crossValue = xAxisCrossing.kind === "min" ? low : xAxisCrossing.kind === "max" ? high : Math.min(Math.max(xAxisCrossing.value, low), high);

            if (crossValue >= high) {
                xAxisPlacement = { kind: "high" };
            } else if (crossValue <= low) {
                xAxisPlacement = { kind: "low" };
            } else {
                hiddenYAxes.push({ type: "value", show: false, min: low - crossValue, max: high - crossValue, inverse: yAxisModel.reversed === true, splitLine: { show: false } });
                xAxisPlacement = { kind: "on-zero" };
            }
        }

        const xAxis = {
            type: "value",
            ...axisStyle,
            ...categoryAxisText,
            // PptxGenJS writes the x axis's numbers in General: both applications print 1000, not ECharts' 1,000 (#616).
            ...(categoryLabelFormat ? {} : { axisLabel: { ...categoryAxisText.axisLabel, formatter: (value: number) => formatChartNumber(value, "General") } }),
            ...(xScale ? { min: xScale.minimum, max: xScale.maximum, interval: xScale.interval } : {}),
            // catAxisOrientation maxMin runs x right to left; a bubble chart's catAxisLabelPos high puts the x labels at
            // the top (shortcut: ECharts moves the axis line and ticks there too, LibreOffice keeps them at the bottom).
            inverse: categoryAxisModel.reversed === true,
            ...(chart.type === "bubble" && categoryAxisModel.labelPosition === "high" ? { position: "top" } : {}),
            ...(xAxisPlacement.kind === "high" ? { position: yAxisModel?.reversed ? "bottom" : "top" } : {}),
            ...(xAxisPlacement.kind === "default" ? {} : { axisLine: { ...axisStyle.axisLine, onZero: xAxisPlacement.kind === "on-zero", ...(xAxisPlacement.kind === "on-zero" ? { onZeroAxisIndex: 1 } : {}) } }),
            splitLine: gridLineOption(categoryAxisModel),
        };
        const firstYAxis = yAxisModel ? valueAxisOption(yAxisModel, 0, yAxisModel.minimum, yAxisModel.maximum, { kind: "default" }) : undefined;
        const yAxis = firstYAxis
            ? {
                  ...firstYAxis,
                  ...(yAxisPlacement.kind === "high" ? { position: categoryAxisModel.reversed ? ("left" as const) : ("right" as const) } : {}),
                  ...(yAxisPlacement.kind === "on-zero" ? { axisLine: { ...firstYAxis.axisLine, onZero: true, onZeroAxisIndex: 1 } } : {}),
              }
            : { type: "value", ...axisStyle, splitLine: { show: false } };

        for (const [seriesIndex, series] of ySeries.entries()) {
            // Each point's label: a scatter point's own label, or its series name and y value (a bubble's size where the
            // target prints it).
            const labels = chart.dataLabels;
            const label = showsPointLabel(labels)
                ? dataLabelOption(labels, POINT_LABEL_POSITIONS[labels.position] ?? "right", dataIndex => {
                      const customLabel = series.labels[dataIndex];
                      const value = labels.valueIsBubbleSize ? (series.sizes?.[dataIndex] ?? 0) : (series.values[dataIndex] ?? 0);

                      return pointLabelText(labels, series.name, value, customLabel === undefined ? { kind: "none" } : { kind: "label", text: customLabel });
                  })
                : undefined;

            // PptxGenJS writes <c:scatterStyle val="lineMarker"/> and each series' <a:ln w="lineSize"> (<a:noFill/> at
            // 0) and <c:smooth>, so Office joins the points in data order: an ECharts line series on a value x axis.
            scatter.push({
                label,
                name: series.name,
                data: series.values.map((value, index) => (chart.type === "bubble" ? [xValues[index], value, series.sizes?.[index] ?? 0] : [xValues[index], value])),
                ...(chart.type === "bubble"
                    ? {
                          type: "scatter",
                          symbolSize: (value: unknown[]) => 8 + Math.sqrt(Number(value[2] ?? 0) / maxBubble) * 32,
                          itemStyle: { opacity: chart.colorOpacity, ...(chart.dataBorder ? { borderColor: chart.dataBorder.color, borderWidth: pointsToPixels(chart.dataBorder.width) } : {}) },
                      }
                    : {
                          type: "line",
                          smooth: chart.lineSmooth,
                          lineStyle: { width: pointsToPixels(chart.lineSize) },
                          symbol: lineSymbol(chart.lineSymbol),
                          symbolSize,
                          // A line series' symbol has no stroke unless borderColor is set: the series color, as on a line chart, unless
                          // lineDataSymbolLineColor sets its own.
                          itemStyle: { opacity: chart.colorOpacity, borderColor: chart.colors[seriesIndex % chart.colors.length], ...symbolBorder },
                      }),
            });
        }

        option = {
            ...common,
            ...(graphics.length ? { graphic: graphics } : {}),
            grid,
            // PptxGenJS writes the x axis (a c:valAx) with catGridLine and the y axis with valGridLine.
            // A hidden axis to cross on is the second of its direction.
            xAxis: hiddenXAxes.length > 0 ? [xAxis, ...hiddenXAxes] : xAxis,
            yAxis: hiddenYAxes.length > 0 ? [yAxis, ...hiddenYAxes] : yAxis,
            series: scatter,
        };
    } else {
        // A date axis draws a category per time-line slot, each series' points in their dates' slots (#580).
        const dateSlots = categoryAxisModel.dateSlots;
        const categories = dateSlots?.labels ?? chart.series[0]?.labels ?? [];
        const slotCategory: number[] = [];

        for (const [category, slot] of (dateSlots?.categorySlots ?? []).entries()) {
            if (slot !== null) {
                slotCategory[slot] = category;
            }
        }

        // Only the slots a major unit labels get a label and a tick.
        const majorSlots = dateSlots?.labels.includes("") ? (index: number) => dateSlots.labels[index] !== "" : undefined;
        const valueAxes: ReturnType<typeof valueAxisOption>[] = [];
        // The value range each axis draws: a percent-stacked axis 0 to 100.
        const ranges: [number, number][] = [];

        for (const [index, axis] of scaledValueAxes.entries()) {
            const percent = chart.series.some(series => series.stack === "percentStacked" && series.valueAxisIndex === index);

            ranges.push(percent ? [0, 100] : [axis.minimum, axis.maximum]);
            valueAxes.push(valueAxisOption(axis, index, percent ? 0 : axis.minimum, percent ? 100 : axis.maximum, horizontal && index === 0 ? { kind: "margin", px: bottomLabelMargin } : { kind: "default" }));
        }

        const [valueAxis] = valueAxes;
        const [low, high] = ranges[0] ?? [0, 0];

        if (!valueAxis || !firstAxis) {
            throw new TypeError(`A ${chart.type} chart model needs at least one value axis; got none`);
        }

        // Where the category axis crosses the value axis (valAxisCrossesAt): its low end, its high end, or a value inside
        // it (autoZero: 0, or the end nearer it). Its line and ticks go there, and unstacked bars and areas start there.
        const crossing = firstAxis.categoryAxisCrossing;
        const crossValue = crossing === undefined ? Math.min(Math.max(0, low), high) : crossing.kind === "min" ? low : crossing.kind === "max" ? high : Math.min(Math.max(crossing.value, low), high);
        const lineAt = crossValue <= low ? "low" : crossValue >= high ? "high" : "inside";
        // The value axis's low end is at the bottom (a horizontal chart's left), or the top (right) when it is reversed.
        const lowSide = horizontal ? (firstAxis.reversed ? "right" : "left") : firstAxis.reversed ? "top" : "bottom";
        const highSide = horizontal ? (firstAxis.reversed ? "left" : "right") : firstAxis.reversed ? "bottom" : "top";
        // The category labels sit by the value axis's low or high end, or with the line (catAxisLabelPos). Away from the
        // line they are a second category axis of their own, as LibreOffice keeps a column chart's labels below its
        // negative values (its default `low`).
        const labelSide = categoryAxisModel.labelPosition === "high" ? highSide : lowSide;
        // shortcut: ECharts keeps an onZero axis's labels at the plot edge, so `nextTo` labels stay at the low end rather
        // than beside a line inside the plot; follow the line once the plot height is known here.
        const labelsApart = (categoryAxisModel.labelPosition === "low" && lineAt !== "low") || (categoryAxisModel.labelPosition === "high" && lineAt !== "high");
        // A crossing inside the axis other than 0 moves the line onto the zero of a hidden copy of the value axis shifted
        // by it, as ECharts puts an axis line on another axis's zero only.
        const shifted = lineAt === "inside" && crossValue !== 0;
        const shiftedIndex = valueAxes.length;
        const categoryAxis = {
            type: "category",
            data: categories,
            ...axisStyle,
            ...categoryAxisText,
            ...(majorSlots ? { axisLabel: { ...categoryAxisText.axisLabel, interval: majorSlots }, axisTick: { ...categoryAxisText.axisTick, interval: majorSlots } } : {}),
            ...(categoryAxisModel.reversed ? { inverse: true } : {}),
            ...(labelsApart ? { axisLabel: { show: false }, name: undefined } : {}),
            // The category axis crosses a reversed value axis at its minimum, now at the far end (LibreOffice draws a
            // maxMin line chart's category labels at the top).
            ...(firstAxis.reversed || lineAt === "high" ? { position: lineAt === "high" ? highSide : lowSide } : {}),
            ...(crossing !== undefined && !shifted ? { axisLine: { ...axisStyle.axisLine, onZero: lineAt === "inside" } } : {}),
            ...(shifted ? { axisLine: { ...axisStyle.axisLine, onZero: true, onZeroAxisIndex: shiftedIndex } } : {}),
            // Category gridlines run through the ticks, between the categories, as Office draws them.
            splitLine: gridLineOption(categoryAxisModel),
            // On the ticks (`<c:crossBetween val="midCat"/>`) a bar is centered on its tick and cut at the plot area's edges.
            // ECharts 6 would widen the axis to hold the bars (containShape), which puts them back between the ticks.
            boundaryGap: !chart.categoriesOnTicks,
            containShape: false,
        } as const;
        const categoryAxes: Record<string, unknown>[] = [categoryAxis];

        if (labelsApart) {
            categoryAxes.push({
                type: "category",
                data: categories,
                position: labelSide,
                ...(categoryAxisModel.reversed ? { inverse: true } : {}),
                ...categoryAxisText,
                axisLine: { show: false, onZero: false },
                axisTick: { show: false },
                splitLine: { show: false },
            });
        }

        // Where the value axis crosses the category axis (catAxisCrossesAt): before the first category, after the last,
        // or before a category in between, the zero of a hidden value axis over the categories. A reversed category axis
        // starts at the other side.
        const valueCrossing = categoryAxisModel.valueAxisCrossing;
        const valueAtEnd = (valueCrossing?.kind === "max") !== (categoryAxisModel.reversed === true);

        if (valueCrossing?.kind === "category") {
            categoryAxes.push({ type: "value", show: false, min: -valueCrossing.index, max: categories.length - valueCrossing.index, inverse: categoryAxisModel.reversed === true, splitLine: { show: false } });
        }

        // Unstacked bars grow from the crossing, and areas fill to it (LibreOffice 26.2 starts a stacked chart's bars at 0).
        const stackedBars = chart.series.some(series => series.type === "bar" && series.stack !== undefined);
        const primaryValueAxis = {
            ...valueAxis,
            ...(valueCrossing?.kind === "category" ? { axisLine: { ...valueAxis.axisLine, onZero: true, onZeroAxisIndex: categoryAxes.length - 1 } } : valueAtEnd ? { position: horizontal ? ("top" as const) : ("right" as const) } : {}),
            ...(crossing !== undefined && !stackedBars ? { startValue: crossValue } : {}),
        };
        // A secondary axis takes the side the primary one leaves: LibreOffice 26.2.6 draws a maxMin category axis's
        // primary value axis on the right and the secondary one on the left (#582). PptxGenJS writes catAxisCrossesAt on
        // both value axes, and LibreOffice draws both lines at that category (#591). With a secondary category axis,
        // catAxisCrossesAt max puts both value axes on the right in PowerPoint and LibreOffice 26.2.6 (#603).
        const secondaryValueAxes =
            valueCrossing?.kind === "category"
                ? valueAxes.slice(1).map(axis => ({ ...axis, axisLine: { ...axis.axisLine, onZero: true, onZeroAxisIndex: categoryAxes.length - 1 } }))
                : valueAtEnd && !horizontal && !(secondaryCategory && valueCrossing?.kind === "max")
                  ? valueAxes.slice(1).map(axis => ({ ...axis, position: "left" as const }))
                  : valueAxes.slice(1);
        const valueAxisList: Record<string, unknown>[] = [primaryValueAxis, ...secondaryValueAxes];

        if (shifted) {
            valueAxisList.push({ type: "value", show: false, min: low - crossValue, max: high - crossValue, inverse: firstAxis.reversed === true, splitLine: { show: false } });
        }

        // The secondary category axis: its labels in their row (below the primary labels, the row's line above them) or
        // over the primary labels, its title in its band; the series on it take its order (onSecondaryCategoryAxis).
        const secondaryAxisIndex = categoryAxes.length;

        if (secondaryCategory) {
            const titleStyle = { color: secondaryCategory.titleColor, fontFamily: secondaryCategory.titleFontFace, fontSize: pointsToPixels(secondaryCategory.titleFontSize) };
            const titled = secondaryTitleBand > 0;
            const rowOffset = secondaryRow && !secondaryHigh && !noBottomLabels ? bottomLabelBand : 0;

            categoryAxes.push({
                type: "category",
                data: categories,
                position: secondaryHigh ? highSide : lowSide,
                offset: rowOffset,
                ...(secondaryCategory.reversed ? { inverse: true } : {}),
                axisLine: secondaryRow && !secondaryHigh && !secondaryCategory.hidden && rowOffset > 0 ? { ...axisStyle.axisLine, onZero: false } : { show: false, onZero: false },
                axisTick: { show: false },
                axisLabel: {
                    show: secondaryLabels,
                    margin: bottomLabelMargin,
                    color: secondaryCategory.labelColor,
                    fontFamily: secondaryCategory.labelFontFace,
                    fontSize: secondaryLabelPx,
                    ...(secondaryCategory.labelBold ? { fontWeight: "bold" } : {}),
                    ...(secondaryCategory.labelItalic ? { fontStyle: "italic" } : {}),
                },
                ...(titled && secondaryRow
                    ? {
                          name: secondaryCategory.title,
                          nameLocation: "middle",
                          nameGap: TICK_LENGTH + AXIS_LABEL_MARGIN + secondaryLabelPx * lineEm + gap + (secondaryHigh ? 0 : categoryTitleBand),
                          nameRotate: secondaryCategory.titleRotate,
                          nameTextStyle: titleStyle,
                      }
                    : {}),
                splitLine: { show: false },
                boundaryGap: !chart.categoriesOnTicks,
            });

            // An overlaid axis's title stands above the plot area, on an axis of its own.
            if (titled && !secondaryRow) {
                categoryAxes.push({
                    type: "category",
                    data: categories,
                    position: highSide,
                    axisLine: { show: false, onZero: false },
                    axisTick: { show: false },
                    axisLabel: { show: false },
                    splitLine: { show: false },
                    name: secondaryCategory.title,
                    nameLocation: "middle",
                    nameGap: (secondaryHigh ? secondaryLabelBand : 0) + gap,
                    nameRotate: secondaryCategory.titleRotate,
                    nameTextStyle: titleStyle,
                });
            }
        }

        // The bar, line and area series on the category axis.
        let barCount = 0;

        for (const series of chart.series) {
            if (series.type === "bar") {
                barCount++;
            }
        }

        // barOverlapPct: each bar of a category overlaps the next by that share of its width (a gap when negative).
        const visibleSeries = stackedBars ? 1 : Math.max(1, barCount);
        const overlap = chart.barOverlapPercent ?? 0;
        const barWidth = `${100 / (visibleSeries - ((visibleSeries - 1) * overlap) / 100 + chart.barGapWidthPercent / 100)}%`;
        const cartesian: Record<string, unknown>[] = [];

        for (const series of chart.series) {
            // A combo chart type's own fill, outline, line and markers, else the chart's (#594).
            const style = series.style ?? chart;
            const border = style.dataBorder ? { borderColor: style.dataBorder.color, borderWidth: pointsToPixels(style.dataBorder.width) } : {};
            const seriesSymbolPixels = pointsToPixels(style.lineSymbolSize);
            const seriesSymbolSize = style.lineSymbol === "dash" ? [seriesSymbolPixels, seriesSymbolPixels * DASH_SYMBOL_HEIGHT_RATIO] : seriesSymbolPixels;
            const seriesSymbolBorder = { borderWidth: pointsToPixels(style.lineSymbolLineSize), ...(style.lineSymbolLineColor ? { borderColor: style.lineSymbolLineColor } : {}) };
            const seriesType = series.type === "bar" ? "bar" : "line";
            const lineLike = series.type === "line" || series.type === "area";
            const color = chart.colors[series.colorIndex % chart.colors.length];
            let values: readonly number[] = series.values;

            // Each value as its percentage (or share of 1) of the category's total over the series of its stack.
            if (series.stack === "percentStacked" || series.stack === "shareStacked") {
                const percent: number[] = [];

                for (const [valueIndex, value] of series.values.entries()) {
                    let total = 0;

                    for (const other of chart.series) {
                        if (other.stack === series.stack && other.type === series.type && other.valueAxisIndex === series.valueAxisIndex) {
                            total += Math.abs(other.values[valueIndex] ?? 0);
                        }
                    }

                    percent.push(total === 0 ? 0 : (value / total) * (series.stack === "percentStacked" ? 100 : 1));
                }

                values = percent;
            }

            // A line or area point's label sits beside its point, a bar's at its end; a combo chart's type has its own.
            const labels = series.dataLabels ?? chart.dataLabels;
            const positions = lineLike ? POINT_LABEL_POSITIONS : horizontal ? BAR_LABEL_POSITIONS : COLUMN_LABEL_POSITIONS;
            const labelPosition = positions[labels.position] ?? "inside";
            // The series' own value, also on a percent-stacked chart that draws its share; on the series color with
            // dataLabelBkgrdColors.
            const label = showsPointLabel(labels)
                ? {
                      ...dataLabelOption(labels, labelPosition, dataIndex => pointLabelText(labels, series.name, series.values[dateSlots ? (slotCategory[dataIndex] ?? -1) : dataIndex] ?? 0, { kind: "none" })),
                      ...(labels.seriesColorBackground ? { backgroundColor: color } : {}),
                  }
                : undefined;
            // A negative bar's label takes its own position, as its end points the other way (#478).
            const negativePosition = label && series.type === "bar" ? ((horizontal ? NEGATIVE_BAR_LABEL_POSITIONS : NEGATIVE_COLUMN_LABEL_POSITIONS)[labels.position] ?? "inside") : labelPosition;
            const pointFill = series.varyColors || series.automaticPointFill;
            let data: readonly unknown[] = values;

            // Each point in its own fill (the series color keys the legend) or with its own label position.
            if (pointFill || (negativePosition !== labelPosition && values.some(value => value < 0))) {
                const points: unknown[] = [];

                for (const [valueIndex, value] of values.entries()) {
                    points.push({
                        value,
                        // A point's own `<c:dPt>` `<c:spPr>` holds its fill and no line, so it drops the dataBorder
                        // (LibreOffice 26.2 draws such a bar without it).
                        ...(pointFill ? { itemStyle: { color: series.pointFills?.[valueIndex] ?? series.automaticPointFill ?? chart.colors[valueIndex % chart.colors.length], ...(series.pointFills ? { borderWidth: 0 } : {}) } } : {}),
                        ...(value < 0 && negativePosition !== labelPosition ? { label: { position: negativePosition } } : {}),
                    });
                }

                data = points;
            }

            // A slot without a date is empty, and LibreOffice joins a line across it (displayBlanksAs span).
            if (dateSlots) {
                const slotted: unknown[] = Array.from({ length: categories.length }, () => "-");

                for (const [category, item] of data.entries()) {
                    const slot = dateSlots.categorySlots[category];

                    if (slot !== null && slot !== undefined) {
                        slotted[slot] = item;
                    }
                }

                data = slotted;
            }

            cartesian.push({
                label,
                name: series.name,
                type: seriesType,
                data,
                ...(dateSlots && lineLike ? { connectNulls: true, showAllSymbol: true } : {}),
                // Each type stacks apart from the others, as each is its own chart element.
                stack: series.stack ? `pptx-${series.type}-${series.valueAxisIndex}` : undefined,
                barWidth: seriesType === "bar" ? barWidth : undefined,
                barGap: seriesType === "bar" ? `${-overlap}%` : undefined,
                // PptxGenJS writes <c:smooth> on line series only, so an area is never smoothed.
                smooth: series.type === "line" ? style.lineSmooth : undefined,
                // PptxGenJS writes no <c:marker> on an area series, and Office draws none; ECharts draws a point's label on
                // its symbol, so a labelled series without markers keeps an invisible one.
                ...(lineLike && label && (series.type === "area" || lineSymbol(style.lineSymbol) === "none")
                    ? { symbol: "circle", symbolSize: 0 }
                    : { symbol: series.type === "line" ? lineSymbol(style.lineSymbol) : series.type === "area" ? "none" : undefined, symbolSize: series.type === "line" ? seriesSymbolSize : undefined }),
                // PptxGenJS writes an area series' <c:spPr> as its chartColors fill at chartColorsOpacity and no <a:ln>
                // (lineSize is only for line and radar series), so LibreOffice fills it opaque by default with no edge;
                // dataBorder gives it one.
                lineStyle:
                    series.type === "line" ? { width: pointsToPixels(style.lineSize), color } : series.type === "area" ? (style.dataBorder ? { width: pointsToPixels(style.dataBorder.width), color: style.dataBorder.color } : { width: 0 }) : undefined,
                areaStyle: series.type === "area" ? { color, opacity: style.colorOpacity, ...(crossing !== undefined ? { origin: crossValue } : {}) } : undefined,
                // A line's markers are filled in markerFill where the target reads PptxGenJS's marker fill (#604).
                itemStyle: { opacity: style.colorOpacity, color: (series.type === "line" ? series.markerFill : undefined) ?? color, ...(series.type === "line" ? { borderColor: color, ...seriesSymbolBorder } : series.type === "bar" ? border : {}) },
                yAxisIndex: series.valueAxisIndex,
                ...(series.onSecondaryCategoryAxis ? { xAxisIndex: secondaryAxisIndex } : {}),
                // PowerPoint and LibreOffice draw a combo chart's area group behind the other types, whatever the call
                // order (#538); ECharts draws a line series (an area's too) at z 3, over the bars at z 2.
                ...(chart.type === "mixed" && series.type === "area" ? { z: 1 } : {}),
            });
        }

        // ECharts places a band's bars from the top down in series order, and draws each series over the ones before it.
        if (reversedSeries || (chart.firstSeriesInFront && chart.type !== "mixed")) {
            cartesian.reverse();
        }

        // A combo chart's area series, drawn behind the others (z 1), swap among themselves only.
        if (chart.firstSeriesInFront && chart.type === "mixed") {
            const slots: number[] = [];
            const areas: Record<string, unknown>[] = [];

            for (const [index, series] of chart.series.entries()) {
                const drawn = cartesian[index];

                if (series.type === "area" && drawn) {
                    slots.push(index);
                    areas.unshift(drawn);
                }
            }

            for (const [order, index] of slots.entries()) {
                cartesian[index] = areas[order] ?? {};
            }
        }

        // A line series without data draws nothing and takes no bar slot; it only names a point's legend entry.
        if (chart.legendListsPoints) {
            for (const entry of legendData ?? []) {
                cartesian.push({ name: entry.name, type: "line", data: [] });
            }
        }

        // The data table under the plot area: the category names, then each series' name (after its key) and values in
        // its data labels' format, each in its category's column, with the borders the target draws.
        if (table) {
            const plotLeft = layout ? layout.x * width : gridLeft;
            const plotRight = layout ? (layout.x + layout.w) * width : width - gridRight;
            const tableTop = (layout ? (layout.y + layout.h) * height : height - gridBottom) + tableOffset;
            const tableLeft = plotLeft - tableNameColumn;
            const tableBottom = tableTop + (chart.series.length + 1) * tableRowPx;
            const column = (plotRight - plotLeft) / Math.max(1, categories.length);
            const text = (value: string, x: number, row: number, align: "center" | "left") => ({
                type: "text",
                x,
                y: tableTop + (row + 0.5) * tableRowPx,
                silent: true,
                style: { text: value, fill: table.color, fontFamily: table.fontFace, fontSize: tablePx, align, verticalAlign: "middle" },
            });
            const rule = (x1: number, y1: number, x2: number, y2: number) => ({ type: "line", silent: true, shape: { x1, y1, x2, y2 }, style: { stroke: table.borderColor, lineWidth: pointsToPixels(table.borderWidth) } });

            for (const [index, category] of categories.entries()) {
                graphics.push(text(categoryLabelFormat && Number.isFinite(Number(category)) && category.trim() !== "" ? formatChartNumber(Number(category), categoryLabelFormat) : category, plotLeft + (index + 0.5) * column, 0, "center"));
            }

            for (const [row, series] of chart.series.entries()) {
                const labels = series.dataLabels ?? chart.dataLabels;
                const nameX = tableLeft + 2 * LEGEND_TEXT_GAP;

                if (table.keys) {
                    graphics.push({ type: "rect", silent: true, shape: { x: nameX, y: tableTop + (row + 1.5) * tableRowPx - tableKeyPx / 2, width: tableKeyPx, height: tableKeyPx }, style: { fill: chart.colors[series.colorIndex % chart.colors.length] } });
                }

                graphics.push(text(series.name, nameX + (table.keys ? tableKeyPx + LEGEND_TEXT_GAP : 0), row + 1, "left"));

                for (const [index, value] of series.values.entries()) {
                    graphics.push(text(formatChartNumber(value, table.valueFormatCode ?? labels.valueFormat), plotLeft + (index + 0.5) * column, row + 1, "center"));
                }
            }

            // The horizontal rules between rows, the vertical ones between columns (the names' first), the outline.
            for (let row = 1; table.horizontalBorder && row <= chart.series.length; row++) {
                graphics.push(rule(tableLeft, tableTop + row * tableRowPx, plotRight, tableTop + row * tableRowPx));
            }

            for (let index = 0; table.verticalBorder && index < categories.length; index++) {
                graphics.push(rule(plotLeft + index * column, index === 0 ? tableTop + tableRowPx : tableTop, plotLeft + index * column, tableBottom));
            }

            if (table.outline) {
                graphics.push({
                    type: "rect",
                    silent: true,
                    shape: { x: tableLeft, y: tableTop + tableRowPx, width: plotRight - tableLeft, height: tableBottom - tableTop - tableRowPx },
                    style: { fill: "none", stroke: table.borderColor, lineWidth: pointsToPixels(table.borderWidth) },
                });
            }
        }

        const categoryAxisOption = categoryAxes.length === 1 ? categoryAxis : categoryAxes;
        const valueAxisOptionList = valueAxisList.length === 1 ? primaryValueAxis : valueAxisList;

        option = {
            ...common,
            grid,
            xAxis: horizontal ? valueAxisOptionList : categoryAxisOption,
            yAxis: horizontal ? categoryAxisOption : valueAxisOptionList,
            series: cartesian,
            ...(graphics.length ? { graphic: graphics } : {}),
        };
    }

    return chartContainer(renderEChartsSvg(width, height, option), style, chart.altText, idScope);
}

/** Resolves caller renderer input to SVG source so the slide model stores only data. */
export function renderChartExtensionSvg(input: ChartExtensionInput, style: NormalizedObjectStyle): string {
    const { width, height } = chartSize(style);

    return input.renderer === "svg" ? input.svg : renderEChartsSvg(width, height, input.option as EChartsCoreOption);
}

export function renderChartExtension(svg: string, style: NormalizedObjectStyle, altText: string, idScope: string): Html {
    return chartContainer(svg, style, altText, idScope);
}
