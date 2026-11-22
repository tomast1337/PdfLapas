import { BarChart, LineChart, PieChart, RadarChart, ScatterChart } from "echarts/charts";
import { GridComponent, LegendComponent, RadarComponent, TitleComponent } from "echarts/components";
import { type EChartsCoreOption, format, init, use } from "echarts/core";
import { SVGRenderer } from "echarts/renderers";
import { automaticAxisInterval, automaticAxisScale } from "#src/chart/axisScale";
import { formatChartNumber } from "#src/chart/format";
import { PPTX_DEFAULTS } from "#src/defaults";
import type { ChartExtensionInput } from "#src/chart/types";
import type { NormalizedChart, NormalizedObjectStyle } from "@pdflapas/model";
import { pointsToPixels } from "@pdflapas/model";
import { renderBar3DSvg } from "./chart3d";
import { type InlineStyle, styleAttribute } from "./css";
import { attributeName, descendants, firstElementChild, getAttribute, parseMarkup, SVG_NAMESPACE, serializeOuter, setAttribute } from "./fragment";
import { element, type Html } from "./markup";
import { applyObjectStyle } from "./style";

use([BarChart, LineChart, PieChart, RadarChart, ScatterChart, GridComponent, LegendComponent, RadarComponent, TitleComponent, SVGRenderer]);

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

/** Tick length in px; a cross tick spans both sides of the axis. */
const TICK_LENGTH = 5;
/** ECharts' default axisLabel.margin, the px between a tick and its label. */
const AXIS_LABEL_MARGIN = 8;
/** ECharts' default axisLabel.textMargin[1], the px it adds on each side of a label's measured width when fitting it. */
const AXIS_LABEL_TEXT_MARGIN = 3;

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

/** ECharts label positions for each OOXML `c:dLblPos`: a column's (vertical), a horizontal bar's, a line point's and a slice's. */
type LabelPositions = Readonly<Partial<Record<NormalizedChart["dataLabels"]["position"], string>>>;
const COLUMN_LABEL_POSITIONS: LabelPositions = { outEnd: "top", inEnd: "insideTop", ctr: "inside", inBase: "insideBottom" };
const BAR_LABEL_POSITIONS: LabelPositions = { outEnd: "right", inEnd: "insideRight", ctr: "inside", inBase: "insideLeft" };
const POINT_LABEL_POSITIONS: LabelPositions = { t: "top", b: "bottom", l: "left", r: "right", ctr: "inside" };

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
const LEGEND_TEXT_GAP = 5;

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
    const titleBand = chart.showTitle ? gap + pointsToPixels(chart.titleFontSize) * pieLayout.titleBandEm + pointsToPixels(plotBand.titleBandPt) : 0;
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
        // Each point of the only series, keyed in its drawn fill, as PowerPoint's PDF export draws it (#546); ECharts keys
        // a bar series' legend by series name, so each entry gets an empty series of its name below.
        // shortcut: a filled <c:dPt> (varyColors) would key its own color; it keeps the series entry until measured (#552).
        const only = chart.series[0];

        legendData = [];

        for (const [index] of (only?.values ?? []).entries()) {
            legendData.push({ name: only?.labels[index] ?? String(index + 1), icon: LEGEND_SQUARE_ICON, itemStyle: { color: only?.automaticPointFill ?? chart.colors[index % chart.colors.length] ?? "#4472C4" } });
        }
    } else if (chart.type !== "radar" && chart.type !== "scatter" && chart.type !== "bubble") {
        legendData = chart.series.map(series => (series.type === "line" ? { name: series.name } : { name: series.name, icon: LEGEND_SQUARE_ICON }));

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
    const common: EChartsCoreOption = {
        animation: false,
        color: chart.colors,
        textStyle: {
            fontFamily: chart.fontFace,
            fontSize: pointsToPixels(chart.fontSize),
        },
        title: chart.showTitle
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
        axisLine: {
            show: chart.axisLineVisible,
            lineStyle: { color: chart.gridLineColor, width: pointsToPixels(chart.gridLineWidth) },
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
    const layout = chart.plotLayout;
    // A manual layout (layoutTarget inner) places the plot rectangle itself, so the labels go where they fall.
    const grid = layout
        ? { left: layout.x * width, top: layout.y * height, width: layout.w * width, height: layout.h * height, outerBoundsMode: "none", ...plotAreaStyle }
        : {
              left: gutter + leftLegendShift,
              // A right legend's band replaces the 2% margin; a secondary axis keeps its gutter beside the band.
              // shortcut: the secondary axis beside a right legend is not measured against LibreOffice; measure it when a
              // deck needs both.
              right: chart.valueAxes.length > 1 ? width * 0.14 + rightLegendBand : Math.max(width * 0.02, rightLegendBand),
              top: gap + titleBand + (chart.showLegend && chart.legendPosition === "t" ? legendBand : 0) + (horizontal ? 0 : valueLabelPx / 2) + topLegendBand,
              bottom: gap + (chart.showLegend && chart.legendPosition === "b" ? legendBand : 0) + categoryTitleBand + bottomLabelMargin + (bottomLabelPx * (plotBand.labelLineEm + 1)) / 2,
              outerBoundsMode: "auto",
              outerBounds: { left: labelBound + leftLegendShift, right: Math.max(labelBound, rightLegendBand), top: 0, bottom: 0 },
              ...plotAreaStyle,
          };
    let option: EChartsCoreOption;

    if (chart.type === "pie" || chart.type === "doughnut") {
        const [series] = chart.series;

        if (!series) {
            throw new TypeError(`A ${chart.type} chart model needs one data series; got none`);
        }

        // Each slice's label: its category, value and share of the total, each shown field on its own line.
        const labels = chart.dataLabels;
        let total = 0;

        for (const value of series.values) {
            total += value;
        }

        const sliceLabel = (index: number) => {
            const value = series.values[index] ?? 0;
            const fields: string[] = [];

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

        const boxWidth = width - left - right;
        const boxHeight = height - top - bottom;
        const radius = Math.max(0, Math.min(boxWidth, boxHeight) / 2);

        option = {
            ...common,
            series: [
                {
                    name: series.name,
                    type: "pie",
                    center: [left + boxWidth / 2, top + boxHeight / 2],
                    // OOXML's c:holeSize is a percentage of the outer diameter.
                    radius: chart.type === "doughnut" ? [(radius * chart.doughnutHoleSize) / 100, radius] : radius,
                    startAngle: 90 - chart.firstSliceAngle,
                    itemStyle: { opacity: chart.colorOpacity },
                    // PptxGenJS 4.0.1 writes a `c:dLbl` for every point with showCatName, showVal and showPercent
                    // off unless showLabel, showValue or showPercent is set, so PowerPoint and LibreOffice draw no
                    // slice labels by default; ECharts would draw category names on leaders. Normalization keeps the
                    // positions inside the slice (bestFit and ctr; LibreOffice places a bestFit label inside).
                    label: labels.showCategory || labels.showValue || labels.showPercent ? dataLabelOption(labels, "inside", sliceLabel) : { show: false },
                    labelLine: { show: false },
                    data: series.values.map((value, index) => ({
                        value,
                        name: series.labels[index] ?? String(index + 1),
                    })),
                },
            ],
        };
    } else if (chart.type === "radar") {
        const categoryCount = Math.max(...chart.series.map(series => Math.max(series.labels.length, series.values.length)));
        const indicators = Array.from({ length: categoryCount }, (_, index) => ({
            name: chart.series[0]?.labels[index] ?? String(index + 1),
            max: Math.max(1, ...chart.series.map(series => Math.abs(series.values[index] ?? 0))) * 1.1,
        }));

        option = {
            ...common,
            radar: { indicator: indicators },
            series: [
                {
                    type: "radar",
                    symbol: chart.radarStyle === "standard" ? "none" : lineSymbol(chart.lineSymbol),
                    symbolSize,
                    itemStyle: symbolBorder,
                    areaStyle: chart.radarStyle === "filled" ? {} : undefined,
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
        // major unit or the same automatic interval.
        // shortcut: the bubble chart's x scale is not measured against LibreOffice (it may leave room for the bubbles);
        // measure it when a deck has one.
        const xScale = xValues.length > 0 ? automaticAxisScale(Math.min(...xValues), Math.max(...xValues)) : undefined;
        const scatter: Record<string, unknown>[] = [];

        for (const [seriesIndex, series] of ySeries.entries()) {
            // PptxGenJS writes <c:scatterStyle val="lineMarker"/> and each series' <a:ln w="lineSize"> (<a:noFill/> at
            // 0) and <c:smooth>, so Office joins the points in data order: an ECharts line series on a value x axis.
            scatter.push({
                name: series.name,
                data: series.values.map((value, index) => (chart.type === "bubble" ? [xValues[index], value, series.sizes?.[index] ?? 0] : [xValues[index], value])),
                ...(chart.type === "bubble"
                    ? { type: "scatter", symbolSize: (value: unknown[]) => 8 + Math.sqrt(Number(value[2] ?? 0) / maxBubble) * 32, itemStyle: { opacity: chart.colorOpacity } }
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
            grid,
            // PptxGenJS writes the x axis (a c:valAx) with catGridLine and the y axis with valGridLine.
            xAxis: {
                type: "value",
                ...axisStyle,
                ...(xScale ? { min: xScale.minimum, max: xScale.maximum, interval: xScale.interval } : {}),
                axisTick: tickOption(chart.categoryMajorTickMark, chart.gridLineColor, chart.gridLineWidth),
                splitLine: gridLineOption(categoryAxisModel),
            },
            yAxis: firstAxis
                ? {
                      type: "value",
                      ...axisStyle,
                      min: firstAxis.minimum,
                      max: firstAxis.maximum,
                      interval: firstAxis.majorUnit ?? automaticAxisInterval(firstAxis.minimum, firstAxis.maximum),
                      axisLabel: {
                          color: firstAxis.labelColor,
                          fontFamily: firstAxis.labelFontFace,
                          fontSize: pointsToPixels(firstAxis.labelFontSize),
                          fontWeight: firstAxis.labelBold ? "bold" : "normal",
                          fontStyle: firstAxis.labelItalic ? "italic" : "normal",
                          formatter: (value: number) => formatChartNumber(value, firstAxis.labelFormatCode),
                      },
                      axisTick: tickOption(firstAxis.majorTickMark, firstAxis.lineColor, firstAxis.lineWidth),
                      splitLine: gridLineOption(firstAxis),
                  }
                : { type: "value", ...axisStyle, splitLine: { show: false } },
            series: scatter,
        };
    } else {
        const categories = chart.series[0]?.labels ?? [];
        const categoryAxis = {
            type: "category",
            data: categories,
            ...axisStyle,
            axisLabel: {
                color: categoryAxisModel.labelColor,
                fontFamily: categoryAxisModel.labelFontFace,
                fontSize: categoryLabelPx,
                ...(horizontal ? {} : { margin: bottomLabelMargin }),
                ...(categoryAxisModel.labelBold ? { fontWeight: "bold" } : {}),
                ...(categoryAxisModel.labelItalic ? { fontStyle: "italic" } : {}),
            },
            axisTick: tickOption(chart.categoryMajorTickMark, chart.gridLineColor, chart.gridLineWidth),
            name: categoryAxisModel.showTitle ? categoryAxisModel.title : undefined,
            nameLocation: "middle",
            nameGap: horizontal ? AXIS_LABEL_MARGIN : TICK_LENGTH + AXIS_LABEL_MARGIN + categoryLabelPx * lineEm + gap,
            nameRotate: categoryAxisModel.titleRotate,
            nameTextStyle: {
                color: categoryAxisModel.titleColor,
                fontFamily: categoryAxisModel.titleFontFace,
                fontSize: pointsToPixels(categoryAxisModel.titleFontSize),
            },
            // Category gridlines run through the ticks, between the categories, as Office draws them.
            splitLine: gridLineOption(categoryAxisModel),
        } as const;
        const valueAxes = [];

        for (const [index, axis] of chart.valueAxes.entries()) {
            const min = chart.grouping === "percentStacked" ? 0 : axis.minimum;
            const max = chart.grouping === "percentStacked" ? 100 : axis.maximum;

            valueAxes.push({
                type: axis.logScaleBase === undefined ? ("value" as const) : ("log" as const),
                logBase: axis.logScaleBase,
                min,
                max,
                interval: axis.logScaleBase === undefined ? (axis.majorUnit ?? automaticAxisInterval(min, max)) : undefined,
                position: index === 1 ? ("right" as const) : ("left" as const),
                show: !axis.hidden,
                name: axis.showTitle ? axis.title : undefined,
                nameLocation: "middle" as const,
                nameGap: 40,
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
                    ...(horizontal && index === 0 ? { margin: bottomLabelMargin } : {}),
                    fontWeight: axis.labelBold ? "bold" : "normal",
                    fontStyle: axis.labelItalic ? "italic" : "normal",
                    formatter: (value: number) => formatChartNumber(value, axis.labelFormatCode),
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
            });
        }

        const [valueAxis] = valueAxes;

        if (!valueAxis) {
            throw new TypeError(`A ${chart.type} chart model needs at least one value axis; got none`);
        }

        // The bar, line and area series on the category axis.
        let barCount = 0;

        for (const series of chart.series) {
            if (series.type === "bar") {
                barCount++;
            }
        }

        const visibleSeries = stacked ? 1 : Math.max(1, barCount);
        const barWidth = `${100 / (visibleSeries + chart.barGapWidthPercent / 100)}%`;
        const cartesian: Record<string, unknown>[] = [];

        for (const series of chart.series) {
            const seriesType = series.type === "bar" ? "bar" : "line";
            const lineLike = series.type === "line" || series.type === "area";
            const color = chart.colors[series.colorIndex % chart.colors.length];
            let values: readonly number[] = series.values;

            // Each value as its percentage of the category's total over the series of its stack.
            if (series.stack === "percentStacked") {
                const percent: number[] = [];

                for (const [valueIndex, value] of series.values.entries()) {
                    let total = 0;

                    for (const other of chart.series) {
                        if (other.stack === series.stack && other.type === series.type && other.valueAxisIndex === series.valueAxisIndex) {
                            total += Math.abs(other.values[valueIndex] ?? 0);
                        }
                    }

                    percent.push(total === 0 ? 0 : (value / total) * 100);
                }

                values = percent;
            }

            // Normalization allows value labels only on bar and line charts, whose positions the tables map.
            // shortcut: a negative column's outEnd label sits above the zero line, not below the bar; flip it per point when needed (#478).
            const positions = series.type === "line" ? POINT_LABEL_POSITIONS : horizontal ? BAR_LABEL_POSITIONS : COLUMN_LABEL_POSITIONS;
            const labelPosition = positions[chart.dataLabels.position] ?? "inside";
            // The series' own value, also on a percent-stacked chart that draws its share.
            const label = chart.dataLabels.showValue ? dataLabelOption(chart.dataLabels, labelPosition, dataIndex => formatChartNumber(series.values[dataIndex] ?? 0, chart.dataLabels.valueFormat)) : undefined;

            cartesian.push({
                label,
                name: series.name,
                type: seriesType,
                // Each point in its own fill; the series color keys the legend.
                data:
                    series.varyColors || series.automaticPointFill
                        ? values.map((value, valueIndex) => ({
                              value,
                              itemStyle: { color: series.automaticPointFill ?? chart.colors[valueIndex % chart.colors.length] },
                          }))
                        : values,
                // Each type stacks apart from the others, as each is its own chart element.
                stack: series.stack ? `pptx-${series.type}-${series.valueAxisIndex}` : undefined,
                barWidth: seriesType === "bar" ? barWidth : undefined,
                barGap: seriesType === "bar" ? "0%" : undefined,
                smooth: lineLike ? chart.lineSmooth : undefined,
                // PptxGenJS writes no <c:marker> on an area series, and Office draws none.
                symbol: series.type === "line" ? lineSymbol(chart.lineSymbol) : series.type === "area" ? "none" : undefined,
                symbolSize: series.type === "line" ? symbolSize : undefined,
                // PptxGenJS writes an area series' <c:spPr> as its chartColors fill at chartColorsOpacity and no <a:ln>
                // (lineSize is only for line and radar series), so LibreOffice fills it opaque by default with no edge.
                lineStyle: series.type === "line" ? { width: pointsToPixels(chart.lineSize), color } : series.type === "area" ? { width: 0 } : undefined,
                areaStyle: series.type === "area" ? { color, opacity: chart.colorOpacity } : undefined,
                itemStyle: { opacity: chart.colorOpacity, color, ...(series.type === "line" ? { borderColor: color, ...symbolBorder } : {}) },
                yAxisIndex: series.valueAxisIndex,
            });
        }

        // ECharts places a band's bars from the top down in series order, and draws each series over the ones before it.
        if (reversedSeries || chart.firstSeriesInFront) {
            cartesian.reverse();
        }

        // A line series without data draws nothing and takes no bar slot; it only names a point's legend entry.
        if (chart.legendListsPoints) {
            for (const entry of legendData ?? []) {
                cartesian.push({ name: entry.name, type: "line", data: [] });
            }
        }

        option = {
            ...common,
            grid,
            xAxis: chart.type === "bar" && chart.barDirection === "bar" ? valueAxis : categoryAxis,
            yAxis: chart.type === "bar" && chart.barDirection === "bar" ? categoryAxis : valueAxes.length === 1 ? valueAxis : valueAxes,
            series: cartesian,
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
