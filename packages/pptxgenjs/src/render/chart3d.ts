import { format } from "echarts/core";
import { PPTX_DEFAULTS } from "#src/defaults";
import { builtInFontMetrics } from "#src/fontMetrics";
import { automaticAxisInterval, type AxisScale, bar3DAxisScale } from "#src/chart/axisScale";
import { pointLabelText, showsPointLabel } from "#src/chart/dataLabel";
import { formatChartNumber } from "#src/chart/format";
import type { NormalizedChart } from "@pdflapas/model";
import { pointsToPixels } from "@pdflapas/model";
import { LEGEND_TEXT_GAP, TICK_LENGTH } from "./chart";
import { type Attribute, element, escapeText, type Html } from "./markup";

type Point = {
    x: number;
    y: number;
};
type Bar = {
    category: number;
    series: number;
    base: number;
    value: number;
};

const n = (value: number): string => Number(value.toFixed(2)).toString();
const p = (point: Point): string => `${n(point.x)},${n(point.y)}`;

function shade(hex: string, amount: number): string {
    const value = hex.replace(/^#/, "");
    const channels = [0, 2, 4].map(offset => Number.parseInt(value.slice(offset, offset + 2), 16));
    const adjusted = channels.map(channel => (amount >= 0 ? Math.round(channel + (255 - channel) * amount) : Math.round(channel * (1 + amount))));

    return `#${adjusted.map(channel => channel.toString(16).padStart(2, "0")).join("")}`;
}

type CameraPoint = {
    /** Distance along the view direction from the box center; larger is farther. */
    distance: number;
} & Point;

const DEGREE = Math.PI / 180;

/**
 * The OOXML view3D camera on a box `depth` and `height` deep and tall per unit of category-axis width: rotY turns
 * it about the vertical axis, rotX tilts it toward the viewer, and the perspective is the field of view in half
 * degrees (ECMA-376 21.2.2.135), which the camera distance turns into for the whole box. Right-angle axes use a
 * parallel projection, as the spec ignores the perspective there.
 * shortcut: Office also keeps the category axis horizontal with right-angle axes; fix when a fixture sets v3DRAngAx.
 */
function view3D(chart: NormalizedChart, depth: number, height: number): (x: number, y: number, z: number) => CameraPoint {
    const yaw = chart.rotationY3D * DEGREE;
    const pitch = chart.rotationX3D * DEGREE;
    const fieldOfView = (chart.perspective3D / 2) * DEGREE;
    const distance = chart.rightAngleAxes3D || fieldOfView === 0 ? Infinity : Math.hypot(1, depth, height) / (2 * Math.tan(fieldOfView / 2));

    return (x, y, z) => {
        const cx = x - 0.5;
        const cy = (y - 0.5) * height;
        const cz = (z - 0.5) * depth;
        const turned = -cx * Math.sin(yaw) + cz * Math.cos(yaw);
        const up = cy * Math.cos(pitch) + turned * Math.sin(pitch);
        const along = -cy * Math.sin(pitch) + turned * Math.cos(pitch);
        const scale = distance === Infinity ? 1 : distance / (distance + along);

        return { x: (cx * Math.cos(yaw) + cz * Math.sin(yaw)) * scale, y: -up * scale, distance: along };
    };
}

const BOX_CORNERS: readonly (readonly [number, number, number])[] = [0, 1].flatMap(x => [0, 1].flatMap(y => [0, 1].map(z => [x, y, z] as const)));

/** Screen bounds of the unit box under a camera. */
function boxBounds(camera: (x: number, y: number, z: number) => CameraPoint) {
    const points = BOX_CORNERS.map(([x, y, z]) => camera(x, y, z));
    const xs = points.map(point => point.x);
    const ys = points.map(point => point.y);

    return { left: Math.min(...xs), top: Math.min(...ys), width: Math.max(...xs) - Math.min(...xs), height: Math.max(...ys) - Math.min(...ys) };
}

/** Deterministic, browser-independent 3D bar projection used instead of WebGL. */
export function renderBar3DSvg(width: number, height: number, chart: NormalizedChart): Html {
    const labels = chart.series[0]?.labels ?? [];
    const categoryCount = Math.max(1, labels.length, ...chart.series.map(series => series.values.length));
    // catAxisOrientation maxMin runs the categories right to left.
    const categorySlot = (category: number): number => (chart.categoryAxis.reversed ? categoryCount - 1 - category : category);
    const seriesCount = Math.max(1, chart.series.length);
    const stacked = chart.grouping === "stacked" || chart.grouping === "percentStacked";
    // The same text bands as the 2D plot area (chart.ts chartGrid): the legend and the category labels below the box.
    const gap = height * PPTX_DEFAULTS.chart.plotAreaGapRatio;
    const lineEm = PPTX_DEFAULTS.text.singleLineEm;
    const legendPx = pointsToPixels(chart.legendFontSize);
    const keySize = Math.round(legendPx * PPTX_DEFAULTS.chart.legendKeyFontRatio);
    // Each series in its color, or each point of the only series keyed as the point is drawn (legendListsPoints):
    // its fill and its outline, as PowerPoint's PDF export draws them (#546).
    const legendEntries: { name: string; fill: string; stroke?: string }[] = [];
    const only = chart.series[0];

    if (chart.legendListsPoints && only) {
        for (const [index] of only.values.entries()) {
            legendEntries.push({ name: only.labels[index] ?? String(index + 1), fill: only.automaticPointFill ?? chart.colors[index % chart.colors.length] ?? "#4472C4", stroke: only.pointLineColors?.[index] });
        }
    } else {
        for (const [index, series] of chart.series.entries()) {
            legendEntries.push({ name: series.name, fill: chart.colors[index % chart.colors.length] ?? "#4472C4" });
        }
    }

    // A left, right or top-right legend lists its entries in a column beside the plot area, a top or bottom one in a row.
    const sideLegend = chart.showLegend && chart.legendPosition !== "b" && chart.legendPosition !== "t";
    let legendColumnWidth = 0;

    if (sideLegend) {
        for (const entry of legendEntries) {
            legendColumnWidth = Math.max(legendColumnWidth, format.getTextRect(entry.name, `${legendPx}px ${chart.legendFontFace}`).width);
        }

        legendColumnWidth += keySize + 5 + 2 * gap;
    }

    const categoryAxisModel = chart.categoryAxis;
    const categoryLabelPx = pointsToPixels(categoryAxisModel.labelFontSize);
    const categoryTitlePx = pointsToPixels(categoryAxisModel.titleFontSize);
    const categoryTitleBand = categoryAxisModel.showTitle ? categoryTitlePx * lineEm + gap : 0;
    // A data table (#592) takes the category labels' band below the box: a row of category names, then one per series.
    const table = chart.dataTable;
    const tablePx = table ? pointsToPixels(table.fontSize) : 0;
    const tableRowPx = tablePx * PPTX_DEFAULTS.chart.dataTable.rowEm;
    const tableBand = table ? (chart.series.length + 1) * tableRowPx + gap : 0;
    // catAxisLabelPos high puts the category labels along the back walls' top edge, as LibreOffice 26.2.6 and PowerPoint
    // draw them (#591).
    const labelsHigh = categoryAxisModel.labelPosition === "high";
    const categoryLabelBand = table ? 0 : categoryLabelPx * lineEm;
    const top = gap + (chart.showTitle ? pointsToPixels(chart.titleFontSize) * lineEm + gap : 0) + (chart.showLegend && chart.legendPosition === "t" ? legendPx * lineEm + gap : 0) + (labelsHigh ? categoryLabelBand : 0);
    const bottom = gap + (chart.showLegend && chart.legendPosition === "b" ? legendPx * lineEm + gap : 0) + (labelsHigh ? 0 : categoryLabelBand) + categoryTitleBand + tableBand;
    const side = width * PPTX_DEFAULTS.chart.bar3DSideMarginRatio;
    const leftBand = chart.showLegend && chart.legendPosition === "l" ? legendColumnWidth : 0;
    const rightBand = sideLegend && chart.legendPosition !== "l" ? legendColumnWidth : 0;
    // Fits the box into the plot area. The depth is fixed (one series slot per category slot); the height is automatic, as
    // in LibreOffice (chart2 BarChart::getPreferredDiagramAspectRatio leaves Y free), so it is chosen to fill the area.
    const boxDepth = (stacked ? 1 : seriesCount) / categoryCount;
    // A manual layout (layoutTarget inner) is the plot area itself.
    const layout = chart.plotLayout;
    const area = layout ? { x: layout.x * width, y: layout.y * height, width: layout.w * width, height: layout.h * height } : { x: side + leftBand, y: top, width: width - 2 * side - leftBand - rightBand, height: height - top - bottom };
    // The box keeps the shape it has without a data table, smaller above it, as PowerPoint draws it (#592).
    const aspect = area.width / Math.max(1e-9, area.height + (layout ? 0 : tableBand));
    // The box's screen aspect falls as it grows taller: bisect the height on a log scale.
    let low = Math.log(0.02);
    let high = Math.log(50);

    for (let step = 0; step < 40; step++) {
        const middle = (low + high) / 2;
        const bounds = boxBounds(view3D(chart, boxDepth, Math.exp(middle)));

        if (bounds.width / Math.max(1e-9, bounds.height) > aspect) {
            low = middle;
        } else {
            high = middle;
        }
    }

    const view = view3D(chart, boxDepth, Math.exp((low + high) / 2));
    const bounds = boxBounds(view);
    const fitScale = Math.min(area.width / bounds.width, area.height / bounds.height);
    const fitLeft = area.x + (area.width - bounds.width * fitScale) / 2 - bounds.left * fitScale;
    const fitTop = area.y + (area.height - bounds.height * fitScale) / 2 - bounds.top * fitScale;
    const camera = (x: number, y: number, z: number): CameraPoint => {
        const point = view(x, y, z);

        return { x: fitLeft + point.x * fitScale, y: fitTop + point.y * fitScale, distance: point.distance };
    };
    // The walls stand on the far sides of the floor, the faces drawn on a bar on the near ones.
    const backDepth = Math.cos(chart.rotationY3D * DEGREE) >= 0 ? 1 : 0;
    const frontDepth = 1 - backDepth;
    const sideX = Math.sin(chart.rotationY3D * DEGREE) >= 0 ? 0 : 1;
    const axis = chart.valueAxes[0];
    const labelFontSize = pointsToPixels(axis?.labelFontSize ?? chart.fontSize);
    const labelFontFace = axis?.labelFontFace ?? chart.fontFace;
    const labelFont = builtInFontMetrics(labelFontFace);
    const labelHeight = labelFontSize * (labelFont.kind === "metrics" ? (labelFont.metrics.winAscent + labelFont.metrics.winDescent) / labelFont.metrics.unitsPerEm : PPTX_DEFAULTS.text.singleLineEm);
    // LibreOffice fits about as many steps on a 3D value axis as its labels' height fits along the axis, at most the
    // automatic scale's increments (26.2.6, 3, 5, 4, 8 in 14pt Arial: 0..8 by 1 on a 136.5pt axis, by 2 on a 95pt one,
    // 0..10 by 5 on a 54pt one); a target may need several label heights per step (labelHeightsPerStep).
    // shortcut: the length is this projection's, which differs from LibreOffice's (#541), and LibreOffice estimates it
    // before its last layout pass (3, 5, 4, 7 in 12pt steps by 1 on an 87.6pt axis); fix with the projection.
    const axisBottom = camera(sideX, 0, frontDepth);
    const axisTop = camera(sideX, 1, frontDepth);
    const axisLength = Math.hypot(axisTop.x - axisBottom.x, axisTop.y - axisBottom.y);
    const stepHeight = labelHeight * (axis?.automaticScale?.labelHeightsPerStep ?? 1);
    const maxIncrements = Math.max(1, Math.min(PPTX_DEFAULTS.chart.automaticScale.maxIncrements, Math.floor(axisLength / stepHeight)));
    let scale: AxisScale;

    if (chart.grouping === "percentStacked") {
        scale = { minimum: 0, maximum: 100, interval: axis?.majorUnit ?? automaticAxisInterval(0, 100, maxIncrements) };
    } else if (axis?.automaticScale) {
        const { low, high, fixedMinimum, fixedMaximum } = axis.automaticScale;

        scale = bar3DAxisScale(low, high, { minimum: fixedMinimum ? axis.minimum : undefined, maximum: fixedMaximum ? axis.maximum : undefined, majorUnit: axis.majorUnit }, maxIncrements);
    } else {
        scale = { minimum: chart.valueAxisMinimum, maximum: chart.valueAxisMaximum, interval: axis?.majorUnit ?? automaticAxisInterval(chart.valueAxisMinimum, chart.valueAxisMaximum, maxIncrements) };
    }

    const { minimum, maximum } = scale;
    const range = Math.max(1e-9, maximum - minimum);
    /** `x` and `depth` run 0..1 along the category axis and the depth axis. */
    const project = (x: number, depth: number, value: number): CameraPoint => camera(x, (value - minimum) / range, depth);
    const polygon = (className: string, points: Point[], fill: string): Html =>
        element("polygon", [
            ["class", className],
            ["points", points.map(p).join(" ")],
            ["fill", fill],
        ]);
    const otherX = 1 - sideX;
    const gridStroke: Attribute[] = [
        ["fill", "none"],
        ["stroke", `#${chart.gridLineColor.replace(/^#/, "")}`],
        ["stroke-width", n(pointsToPixels(chart.gridLineWidth))],
    ];
    // The value gridlines take valGridLine (LibreOffice draws them on the walls in its color and width, none with
    // style none); the wall frame keeps the chart's line.
    const valueGridStroke: Attribute[] = axis
        ? [
              ["fill", "none"],
              ["stroke", axis.gridLineColor],
              ["stroke-width", n(pointsToPixels(axis.gridLineWidth))],
              ["stroke-dasharray", axis.gridLineStyle === "dash" ? n(pointsToPixels(axis.gridLineWidth) * 4) : axis.gridLineStyle === "dot" ? n(pointsToPixels(axis.gridLineWidth)) : undefined],
          ]
        : gridStroke;
    const drawsValueGrid = axis?.gridLineStyle !== "none";
    // valAxisHidden and valAxisLabelPos none draw no labels; high puts them at the far edge of the back wall.
    const drawsValueLabels = axis !== undefined && !axis.hidden && axis.labelPosition !== "none";
    // catAxisCrossesAt puts the value axis before a category (a reversed axis counts from the right) or after the last,
    // as LibreOffice 26.2.6 and PowerPoint draw it (#591); valAxisLabelPos high at the far edge.
    const valueCrossing = categoryAxisModel.valueAxisCrossing;
    const crossingX = valueCrossing?.kind === "category" ? valueCrossing.index / categoryCount : valueCrossing?.kind === "max" ? 1 : undefined;
    const labelX0 = crossingX !== undefined ? (categoryAxisModel.reversed ? 1 - crossingX : crossingX) : axis?.labelPosition === "high" ? 1 - sideX : sideX;
    const labelDepth = axis?.labelPosition === "high" ? backDepth : frontDepth;
    // The value labels take the value axis's font.
    const labelStyle: Attribute[] = [
        ["font-family", labelFontFace],
        ["font-size", n(labelFontSize)],
    ];

    if (axis) {
        labelStyle.push(["fill", axis.labelColor]);
    }

    if (axis?.labelBold) {
        labelStyle.push(["font-weight", "bold"]);
    }

    if (axis?.labelItalic) {
        labelStyle.push(["font-style", "italic"]);
    }

    // Office draws the value gridlines across the side wall and the back wall, labelled at the wall's front edge, at
    // each whole step from the minimum (toPrecision drops the floating point noise of 0.9 / 0.1).
    const grid: Html[] = [];
    const tickCount = Math.floor(Number((range / scale.interval).toPrecision(12)));

    // Value labels at a catAxisCrossesAt crossing stand in front of the bars, as both applications draw them (#591).
    const crossingLabels: Html[] = [];
    let labelEdge = labelX0 < 1 ? Infinity : -Infinity;

    for (let index = 0; index <= tickCount; index++) {
        const value = Number((minimum + index * scale.interval).toPrecision(12));

        if (drawsValueGrid) {
            const line = [project(sideX, frontDepth, value), project(sideX, backDepth, value), project(otherX, backDepth, value)];

            grid.push(element("polyline", [["points", line.map(p).join(" ")], ...valueGridStroke]));
        }

        if (drawsValueLabels) {
            const label = project(labelX0, labelDepth, value);
            const x = labelX0 < 1 ? label.x - 8 : label.x + 8;
            const y = label.y + labelFontSize / 4;
            const rotate: Attribute[] = axis.labelRotate ? [["transform", `rotate(${n(axis.labelRotate)} ${n(x)} ${n(y)})`]] : [];
            const tickLabel = formatChartNumber(axis.displayUnit ? value / axis.displayUnit.divisor : value, axis.labelFormatCode);
            const labelWidth = format.getTextRect(tickLabel, `${labelFontSize}px ${labelFontFace}`).width;

            labelEdge = labelX0 < 1 ? Math.min(labelEdge, x - labelWidth) : Math.max(labelEdge, x + labelWidth);
            (crossingX === undefined ? grid : crossingLabels).push(element("text", [["x", n(x)], ...(labelX0 < 1 ? [["text-anchor", "end"] as Attribute] : []), ["y", n(y)], ...rotate, ...labelStyle], escapeText(tickLabel)));
        }
    }

    // The value axis line along the labelled wall edge, and its tick marks at each step (major) and fifth of a step
    // (minor): outside toward the labels, inside away from them, or across (normalizeChart keeps them only where the
    // target draws them, drawsBar3DValueAxisLine, #565).
    if (axis && !axis.hidden) {
        const axisStroke: Attribute[] = [
            ["stroke", axis.lineColor],
            ["stroke-width", n(pointsToPixels(axis.lineWidth))],
            ["stroke-dasharray", axis.lineStyle === "dash" ? n(pointsToPixels(axis.lineWidth) * 4) : axis.lineStyle === "dot" ? n(pointsToPixels(axis.lineWidth)) : undefined],
        ];
        const outward = labelX0 < 1 ? -1 : 1;

        if (axis.lineVisible) {
            const start = project(labelX0, labelDepth, minimum);
            const end = project(labelX0, labelDepth, maximum);

            grid.push(element("line", [["class", "bar3d-value-axis"], ["x1", n(start.x)], ["y1", n(start.y)], ["x2", n(end.x)], ["y2", n(end.y)], ...axisStroke]));
        }

        for (const [mark, steps] of [
            [axis.majorTickMark, tickCount],
            [axis.minorTickMark, tickCount * 5],
        ] as const) {
            for (let index = 0; mark !== "none" && index <= steps; index++) {
                const point = project(labelX0, labelDepth, minimum + (index * tickCount * scale.interval) / steps);
                const from = mark === "inside" ? 0 : TICK_LENGTH;
                const to = mark === "outside" ? 0 : -TICK_LENGTH;

                grid.push(element("line", [["class", "bar3d-value-tick"], ["x1", n(point.x + outward * from)], ["y1", n(point.y)], ["x2", n(point.x + outward * to)], ["y2", n(point.y)], ...axisStroke]));
            }
        }
    }

    // The value axis title, turned up, beside its labels (LibreOffice draws it left of the plot area).
    if (axis?.showTitle) {
        const titlePx = pointsToPixels(axis.titleFontSize);
        const middle = (project(sideX, frontDepth, minimum).y + project(sideX, frontDepth, maximum).y) / 2;
        const edge = Number.isFinite(labelEdge) ? labelEdge : project(labelX0, labelDepth, minimum).x;
        // Office prints it whole, so where the labels leave too little side margin it stays inside the chart's edge, beside
        // the middle steps' labels, which are narrower than the widest one (#568).
        const halfLine = (titlePx * lineEm) / 2;
        const x = labelX0 < 1 ? Math.max(halfLine, edge - gap - titlePx / 2) : Math.min(width - halfLine, edge + gap + titlePx / 2);
        const rotation = axis.titleRotate ?? -90;

        grid.push(
            element(
                "text",
                [
                    ["class", "bar3d-value-title"],
                    ["x", n(x)],
                    ["y", n(middle)],
                    ["text-anchor", "middle"],
                    ["dominant-baseline", "central"],
                    ["transform", `rotate(${n(rotation)} ${n(x)} ${n(middle)})`],
                    ["font-family", axis.titleFontFace],
                    ["font-size", n(titlePx)],
                    ["fill", axis.titleColor],
                ],
                escapeText(axis.title),
            ),
        );
    }

    // The category gridlines (catGridLine) at each category boundary across the floor and up the back wall.
    if (categoryAxisModel.gridLineStyle !== "none") {
        const categoryStroke: Attribute[] = [
            ["fill", "none"],
            ["stroke", categoryAxisModel.gridLineColor],
            ["stroke-width", n(pointsToPixels(categoryAxisModel.gridLineWidth))],
        ];

        for (let index = 0; index <= categoryCount; index++) {
            const x = index / categoryCount;
            const line = [project(x, frontDepth, minimum), project(x, backDepth, minimum), project(x, backDepth, maximum)];

            grid.push(element("polyline", [["class", "bar3d-category-grid"], ["points", line.map(p).join(" ")], ...categoryStroke]));
        }
    }

    // The series gridlines (serGridLine) at each series slot's edge: across the floor, then up the side wall, as
    // LibreOffice 26.2.6 draws them (#581).
    const seriesGrid = chart.seriesAxis;

    if (seriesGrid && seriesGrid.gridLineStyle !== "none") {
        const seriesStroke: Attribute[] = [
            ["fill", "none"],
            ["stroke", seriesGrid.gridLineColor],
            ["stroke-width", n(pointsToPixels(seriesGrid.gridLineWidth))],
            ["stroke-dasharray", seriesGrid.gridLineStyle === "dash" ? n(pointsToPixels(seriesGrid.gridLineWidth) * 4) : seriesGrid.gridLineStyle === "dot" ? n(pointsToPixels(seriesGrid.gridLineWidth)) : undefined],
        ];
        const slots = stacked ? 1 : seriesCount;

        for (let index = 0; index <= slots; index++) {
            const depth = index / slots;
            const line = [project(otherX, depth, minimum), project(sideX, depth, minimum), project(sideX, depth, maximum)];

            grid.push(element("polyline", [["class", "bar3d-series-grid"], ["points", line.map(p).join(" ")], ...seriesStroke]));
        }
    }

    const backSideBottom = project(sideX, backDepth, minimum);
    const backSideTop = project(sideX, backDepth, maximum);
    const backOtherBottom = project(otherX, backDepth, minimum);
    const backOtherTop = project(otherX, backDepth, maximum);
    const frontSideBottom = project(sideX, frontDepth, minimum);
    const frontSideTop = project(sideX, frontDepth, maximum);
    const frontOtherBottom = project(otherX, frontDepth, minimum);
    const floor = polygon("bar3d-floor", [frontSideBottom, frontOtherBottom, backOtherBottom, backSideBottom], `#${PPTX_DEFAULTS.chart.bar3DLighting.floorColor}`);
    const wall = [
        polygon("bar3d-wall", [backSideBottom, backOtherBottom, backOtherTop, backSideTop], `#${PPTX_DEFAULTS.chart.bar3DLighting.wallColor}`),
        polygon("bar3d-wall", [frontSideBottom, backSideBottom, backSideTop, frontSideTop], `#${PPTX_DEFAULTS.chart.bar3DLighting.wallColor}`),
    ];
    const frame = [
        element("polyline", [["class", "bar3d-frame"], ["points", [frontSideTop, frontSideBottom, frontOtherBottom, backOtherBottom, backOtherTop, backSideTop, frontSideTop].map(p).join(" ")], ...gridStroke]),
        element("polyline", [["class", "bar3d-frame"], ["points", [backSideBottom, backSideTop].map(p).join(" ")], ...gridStroke]),
    ];
    const gradientStop = (offset: string, color: string): Html =>
        element("stop", [
            ["offset", offset],
            ["stop-color", color],
        ]);
    // A front gradient per chart color, and one per series whose points take their automatic fill.
    const fronts: [string, string][] = chart.colors.map((color, index) => [`bar3d-front-${index}`, color]);

    for (const [index, series] of chart.series.entries()) {
        if (series.automaticPointFill) {
            fronts.push([`bar3d-front-point-${index}`, series.automaticPointFill]);
        }
    }

    const gradients = element(
        "defs",
        [],
        fronts.map(([id, color]) =>
            element(
                "linearGradient",
                [
                    ["id", id],
                    ["x1", "0"],
                    ["y1", "0"],
                    ["x2", "0"],
                    ["y2", "1"],
                ],
                [gradientStop("0", shade(color, 0.13)), gradientStop("1", shade(color, -0.08))],
            ),
        ),
    );

    const bars: Bar[] = [];

    if (chart.grouping === "stacked" || chart.grouping === "percentStacked") {
        for (let category = 0; category < categoryCount; category++) {
            let positive = 0;
            let negative = 0;
            let total = 0;

            for (const series of chart.series) {
                total += Math.abs(series.values[category] ?? 0);
            }

            for (const [seriesIndex, series] of chart.series.entries()) {
                let value = series.values[category] ?? 0;

                if (chart.grouping === "percentStacked") {
                    value = total === 0 ? 0 : (value / total) * 100;
                }

                const base = value >= 0 ? positive : negative;

                bars.push({ category, series: seriesIndex, base, value: base + value });

                if (value >= 0) {
                    positive += value;
                } else {
                    negative += value;
                }
            }
        }
    } else {
        // valAxisCrossesAt: unstacked bars grow from the crossing (the axis's minimum, maximum or a value inside it), as
        // LibreOffice 26.2.6 and PowerPoint draw them (#591).
        const crossing = axis?.categoryAxisCrossing;
        const base = crossing === undefined ? 0 : crossing.kind === "min" ? minimum : crossing.kind === "max" ? maximum : Math.max(minimum, Math.min(maximum, crossing.value));

        for (let series = seriesCount - 1; series >= 0; series--) {
            for (let category = 0; category < categoryCount; category++) {
                bars.push({ category, series, base, value: chart.series[series]?.values[category] ?? 0 });
            }
        }
    }

    // gapWidth and gapDepth are percentages of the bar's width and depth (c:gapWidth, c:gapDepth); each category
    // has one slot along the category axis and, unless stacked, each series one slot along the depth axis.
    const halfWidth = 0.5 / (1 + chart.barGapWidthPercent / 100) / categoryCount;
    const depthSlot = stacked ? 1 : 1 / seriesCount;
    const depthThickness = depthSlot / (1 + chart.barGapDepthPercent / 100);
    // serAxisOrientation maxMin puts the first series at the back.
    const slotDepth = (series: number): number => (stacked ? 0.5 : ((chart.seriesAxis?.reversed ? seriesCount - 1 - series : series) + 0.5) * depthSlot);
    const barCenter = (bar: Bar): CameraPoint => project((categorySlot(bar.category) + 0.5) / categoryCount, slotDepth(bar.series), (bar.base + bar.value) / 2);
    // Painter's order: the bar farthest from the camera first.
    const ordered: { bar: Bar; distance: number }[] = [];

    for (const bar of bars) {
        ordered.push({ bar, distance: barCenter(bar).distance });
    }

    ordered.sort((a, b) => b.distance - a.distance);

    const barSvg: Html[] = [];

    for (const { bar } of ordered) {
        const cx = (categorySlot(bar.category) + 0.5) / categoryCount;
        const cz = slotDepth(bar.series);
        const x0 = cx - halfWidth;
        const x1 = cx + halfWidth;
        const z0 = Math.max(0, cz - depthThickness / 2);
        const z1 = Math.min(1, cz + depthThickness / 2);
        const base = Math.max(minimum, Math.min(maximum, bar.base));
        const value = Math.max(minimum, Math.min(maximum, bar.value));
        const low = Math.min(base, value);
        const high = Math.max(base, value);
        const series = chart.series[bar.series];
        const color = series?.automaticPointFill ?? chart.colors[bar.series % chart.colors.length] ?? "#4472C4";
        const frontColor = series?.automaticPointFill ? `url(#bar3d-front-point-${bar.series})` : `url(#bar3d-front-${bar.series % chart.colors.length})`;
        // A point's own outline (its `<c:dPt>` `<a:ln>`, no width), or the series' dataBorder.
        const pointOutline = series?.pointLineColors?.[bar.category];
        const outline = pointOutline ?? chart.dataBorder?.color;
        const outlineWidth = pointOutline || !chart.dataBorder ? PPTX_DEFAULTS.chart.pointLineWidthPt : chart.dataBorder.width;
        const topColor = shade(color, PPTX_DEFAULTS.chart.bar3DLighting.topLighten);
        const sideColor = shade(color, -PPTX_DEFAULTS.chart.bar3DLighting.sideDarken);
        const opacity = n(chart.colorOpacity);
        const bfl = project(x0, z0, low);
        const bfr = project(x1, z0, low);
        const bbr = project(x1, z1, low);
        const bbl = project(x0, z1, low);
        const nearZ = backDepth === 1 ? z0 : z1;
        const nearX = sideX === 0 ? x1 : x0;
        const cap = chart.rotationX3D >= 0 ? high : low;
        const front = [project(x0, nearZ, low), project(x1, nearZ, low), project(x1, nearZ, high), project(x0, nearZ, high)];
        const side = [project(nearX, z0, low), project(nearX, z1, low), project(nearX, z1, high), project(nearX, z0, high)];
        const group: Attribute[] = [
            ["class", `bar3d-mark bar3d-${chart.bar3DShape}`],
            ["data-series", String(bar.series)],
            ["data-category", String(bar.category)],
            ["opacity", opacity],
        ];

        // The point's outline along every edge of its faces.
        if (outline) {
            group.push(["stroke", outline], ["stroke-width", n(pointsToPixels(outlineWidth))], ["stroke-linejoin", "round"]);
        }

        if (chart.bar3DShape === "pyramid" || chart.bar3DShape === "pyramidToMax") {
            const apex = project(cx, (z0 + z1) / 2, high);

            barSvg.push(element("g", group, [polygon("bar3d-front", [bfl, bfr, apex], frontColor), polygon("bar3d-side", [bfr, bbr, apex], sideColor), polygon("bar3d-back", [bbr, bbl, apex], topColor)]));
            continue;
        }

        if (chart.bar3DShape === "cone" || chart.bar3DShape === "coneToMax") {
            const bottom = project(cx, (z0 + z1) / 2, low);
            const apex = project(cx, (z0 + z1) / 2, high);
            const radiusX = Math.max(3, Math.hypot(bfr.x - bfl.x, bfr.y - bfl.y) / 2);
            const radiusY = Math.max(2, Math.abs(bbr.y - bfr.y) / 2);

            barSvg.push(
                element("g", group, [
                    element("path", [
                        ["class", "bar3d-front"],
                        ["d", `M ${n(bottom.x - radiusX)} ${n(bottom.y)} Q ${n(bottom.x)} ${n(bottom.y + radiusY)} ${n(bottom.x + radiusX)} ${n(bottom.y)} L ${n(apex.x)} ${n(apex.y)} Z`],
                        ["fill", frontColor],
                    ]),
                    element("ellipse", [
                        ["class", "bar3d-base"],
                        ["cx", n(bottom.x)],
                        ["cy", n(bottom.y)],
                        ["rx", n(radiusX)],
                        ["ry", n(radiusY)],
                        ["fill", sideColor],
                    ]),
                ]),
            );
            continue;
        }

        if (chart.bar3DShape === "cylinder") {
            const bottom = project(cx, (z0 + z1) / 2, low);
            const top = project(cx, (z0 + z1) / 2, high);
            const radiusX = Math.max(3, Math.hypot(bfr.x - bfl.x, bfr.y - bfl.y) / 2);
            const radiusY = Math.max(2, Math.abs(bbr.y - bfr.y) / 2);

            barSvg.push(
                element("g", group, [
                    element("path", [
                        ["class", "bar3d-front"],
                        ["d", `M ${n(top.x - radiusX)} ${n(top.y)} L ${n(bottom.x - radiusX)} ${n(bottom.y)} A ${n(radiusX)} ${n(radiusY)} 0 0 0 ${n(bottom.x + radiusX)} ${n(bottom.y)} L ${n(top.x + radiusX)} ${n(top.y)} Z`],
                        ["fill", frontColor],
                    ]),
                    element("ellipse", [
                        ["class", "bar3d-top"],
                        ["cx", n(top.x)],
                        ["cy", n(top.y)],
                        ["rx", n(radiusX)],
                        ["ry", n(radiusY)],
                        ["fill", topColor],
                    ]),
                ]),
            );
            continue;
        }

        barSvg.push(element("g", group, [polygon("bar3d-front", front, frontColor), polygon("bar3d-side", side, sideColor), polygon("bar3d-top", [project(x0, z0, cap), project(x1, z0, cap), project(x1, z1, cap), project(x0, z1, cap)], topColor)]));
    }

    // The data labels above each bar's top (LibreOffice's bar3D labels), in the order the bars are painted.
    const dataLabelSvg: Html[] = [];
    const dataLabels = chart.dataLabels;

    if (showsPointLabel(dataLabels)) {
        const labelPx = pointsToPixels(dataLabels.fontSize);

        for (const { bar } of ordered) {
            const series = chart.series[bar.series];
            const value = series?.values[bar.category] ?? 0;
            const anchor = project((categorySlot(bar.category) + 0.5) / categoryCount, slotDepth(bar.series), Math.max(minimum, Math.min(maximum, Math.max(bar.base, bar.value))));

            dataLabelSvg.push(
                element(
                    "text",
                    [
                        ["class", "bar3d-data-label"],
                        ["x", n(anchor.x)],
                        ["y", n(anchor.y - labelPx / 2)],
                        ["text-anchor", "middle"],
                        ["font-family", dataLabels.fontFace],
                        ["font-size", n(labelPx)],
                        ["fill", dataLabels.color],
                        ["font-weight", dataLabels.bold ? "bold" : undefined],
                        ["font-style", dataLabels.italic ? "italic" : undefined],
                    ],
                    escapeText(pointLabelText(dataLabels, series?.name ?? "", value, { kind: "none" })),
                ),
            );
        }
    }

    // The category labels and title take the category axis's text style.
    const categoryLabelStyle: Attribute[] = [
        ["font-family", categoryAxisModel.labelFontFace],
        ["font-size", n(categoryLabelPx)],
        ["fill", categoryAxisModel.labelColor],
        ["font-weight", categoryAxisModel.labelBold ? "bold" : undefined],
        ["font-style", categoryAxisModel.labelItalic ? "italic" : undefined],
    ];
    let categoryBottom = 0;
    const categoryLabels: Html[] = [];

    // A deleted axis (catAxisHidden) draws no labels or title, `none` no labels; catAxisLabelRotate turns each label
    // clockwise about its anchor. A data table's header row names the categories instead, as LibreOffice 26.2.6 draws it.
    // shortcut: PowerPoint keeps the labels on the floor's edge above its table too (#592); add a target flag when a
    // PowerPoint deck needs them twice.
    for (let index = 0; !table && !categoryAxisModel.hidden && categoryAxisModel.labelPosition !== "none" && index < categoryCount; index++) {
        // catAxisLabelFrequency (where the target reads it) labels every nth category; catLabelFormatCode formats a
        // numeric one.
        if (categoryAxisModel.labelSkip && index % categoryAxisModel.labelSkip !== 0) {
            continue;
        }

        const edge = project((categorySlot(index) + 0.5) / categoryCount, labelsHigh ? backDepth : frontDepth, labelsHigh ? maximum : minimum);
        // A high label sits above the edge: its anchor 20px lower is the same distance below a low label's.
        const point = labelsHigh ? { ...edge, y: edge.y - 20 - categoryLabelPx * 0.6 } : edge;
        const text = labels[index] ?? String(index + 1);
        const number = Number(text);
        const label = categoryAxisModel.labelFormatCode && text.trim() !== "" && Number.isFinite(number) ? formatChartNumber(number, categoryAxisModel.labelFormatCode) : text;
        const rotate: Attribute[] = categoryAxisModel.labelRotate ? [["transform", `rotate(${n(categoryAxisModel.labelRotate)} ${n(point.x)} ${n(point.y + 20)})`]] : [];

        categoryBottom = Math.max(categoryBottom, point.y + 20);
        categoryLabels.push(element("text", [["x", n(point.x)], ["y", n(point.y + 20)], ["text-anchor", "middle"], ...rotate, ...categoryLabelStyle], escapeText(label)));
    }

    if (categoryAxisModel.showTitle && !categoryAxisModel.hidden) {
        const left = project(0, frontDepth, minimum);
        const right = project(1, frontDepth, minimum);

        categoryLabels.push(
            element(
                "text",
                [
                    ["class", "bar3d-category-title"],
                    ["x", n((left.x + right.x) / 2)],
                    ["y", n(categoryBottom + gap + categoryTitlePx)],
                    ["text-anchor", "middle"],
                    ["font-family", categoryAxisModel.titleFontFace],
                    ["font-size", n(categoryTitlePx)],
                    ["fill", categoryAxisModel.titleColor],
                    ["transform", categoryAxisModel.titleRotate ? `rotate(${n(categoryAxisModel.titleRotate)} ${n((left.x + right.x) / 2)} ${n(categoryBottom + gap + categoryTitlePx)})` : undefined],
                ],
                escapeText(categoryAxisModel.title),
            ),
        );
    }

    // LibreOffice 26.2.6 names the depth slots only when there are several: a single series or a stacked chart, whose
    // series share one slot, gets no depth-axis label. A target may name a single unstacked series (namesSingleDepthSlot).
    // A deleted depth axis (serAxisHidden) names none; its labels and title take the serAxis* styles.
    const seriesAxis = chart.seriesAxis;
    const namesSlots = !stacked && !seriesAxis?.hidden && (chart.series.length > 1 || chart.namesSingleDepthSlot === true);
    const seriesLabelStyle: Attribute[] = seriesAxis
        ? [
              ["font-family", seriesAxis.labelFontFace],
              ["font-size", n(pointsToPixels(seriesAxis.labelFontSize))],
              ["fill", seriesAxis.labelColor],
              ["font-weight", seriesAxis.labelBold ? "bold" : undefined],
              ["font-style", seriesAxis.labelItalic ? "italic" : undefined],
          ]
        : [];
    const seriesLabels: Html[] = [];

    for (const [index, series] of (namesSlots ? chart.series : []).entries()) {
        // serAxisLabelFrequency (where the target reads it) names every nth series.
        if (seriesAxis?.labelSkip && index % seriesAxis.labelSkip !== 0) {
            continue;
        }

        const point = project(otherX, slotDepth(index), minimum);

        seriesLabels.push(element("text", [["class", "bar3d-series-label"], ["x", n(point.x + 7)], ["y", n(point.y + 12)], ...seriesLabelStyle], escapeText(series.name)));
    }

    // showSerAxisTitle: LibreOffice 26.2 prints it below and right of the front slot's name.
    if (seriesAxis?.showTitle && !seriesAxis.hidden) {
        const front = project(otherX, depthSlot / 2, minimum);
        const titlePx = pointsToPixels(seriesAxis.titleFontSize);
        let nameWidth = 0;

        for (const series of namesSlots ? chart.series : []) {
            nameWidth = Math.max(nameWidth, format.getTextRect(series.name, `${pointsToPixels(seriesAxis.labelFontSize)}px ${seriesAxis.labelFontFace}`).width);
        }

        const end = { x: front.x + nameWidth, y: front.y };
        const y = end.y + 12 + titlePx * lineEm;

        seriesLabels.push(
            element(
                "text",
                [
                    ["class", "bar3d-series-title"],
                    ["x", n(end.x + 7)],
                    ["y", n(y)],
                    ["font-family", seriesAxis.titleFontFace],
                    ["font-size", n(titlePx)],
                    ["fill", seriesAxis.titleColor],
                    ["transform", seriesAxis.titleRotate ? `rotate(${n(seriesAxis.titleRotate)} ${n(end.x + 7)} ${n(y)})` : undefined],
                ],
                escapeText(seriesAxis.title),
            ),
        );
    }

    // The legend has its own font (legendFontFace/legendFontSize/legendColor), not the axis-label font of the <g>.
    const legendFont: Attribute[] = [
        ["font-family", chart.legendFontFace],
        ["font-size", n(pointsToPixels(chart.legendFontSize))],
        ["fill", chart.legendColor],
    ];
    const legend: Html[] = [];

    if (chart.showLegend) {
        const pitch = legendPx * lineEm;
        const columnTop = chart.legendPosition === "tr" ? top : height / 2 - (pitch * legendEntries.length) / 2 + pitch / 2;

        for (const [index, { name, fill, stroke }] of legendEntries.entries()) {
            const itemWidth = Math.min(100, width / Math.max(2, legendEntries.length + 1));
            // A row centered below or above the plot area, or a column centered beside it (top-right: at its top).
            const x = !sideLegend ? width / 2 + (index - (legendEntries.length - 1) / 2) * itemWidth : chart.legendPosition === "l" ? gap + keySize + 22 : width - legendColumnWidth + gap + keySize + 22;
            const y = !sideLegend ? (chart.legendPosition === "t" ? top - gap : height - 12) : columnTop + index * pitch;

            legend.push(
                element("rect", [
                    ["x", n(x - 22 - keySize)],
                    ["y", n(y - keySize)],
                    ["width", n(keySize)],
                    ["height", n(keySize)],
                    ["fill", fill],
                    ["stroke", stroke],
                    ["stroke-width", stroke ? n(pointsToPixels(PPTX_DEFAULTS.chart.pointLineWidthPt)) : undefined],
                ]),
                element("text", [["class", "bar3d-legend"], ["x", n(x - 17)], ["y", n(y)], ...legendFont], escapeText(name)),
            );
        }
    }

    // The data table under the box (#592): the category names, then each series' name (after its key) and values in its
    // data labels' format, one column per category across the chart, with the borders the model asks for.
    const tableSvg: Html[] = [];

    if (table) {
        const keyPx = tablePx * PPTX_DEFAULTS.chart.dataTable.keyEm;
        let widestName = 0;

        for (const series of chart.series) {
            widestName = Math.max(widestName, format.getTextRect(series.name, `${tablePx}px ${table.fontFace}`).width);
        }

        const tableLeft = side;
        const tableRight = width - side;
        const columnsLeft = tableLeft + 3 * LEGEND_TEXT_GAP + (table.keys ? keyPx + LEGEND_TEXT_GAP : 0) + widestName;
        const tableTop = area.y + area.height + gap;
        const tableBottom = tableTop + (chart.series.length + 1) * tableRowPx;
        const column = (tableRight - columnsLeft) / categoryCount;
        const textStyle: Attribute[] = [
            ["font-family", table.fontFace],
            ["font-size", n(tablePx)],
            ["fill", table.color],
            ["dominant-baseline", "central"],
        ];
        const rule = (x1: number, y1: number, x2: number, y2: number): Html =>
            element("line", [
                ["class", "bar3d-table-rule"],
                ["x1", n(x1)],
                ["y1", n(y1)],
                ["x2", n(x2)],
                ["y2", n(y2)],
                ["stroke", table.borderColor],
                ["stroke-width", n(pointsToPixels(table.borderWidth))],
            ]);

        for (let index = 0; index < categoryCount; index++) {
            const text = labels[index] ?? String(index + 1);
            const number = Number(text);
            const label = categoryAxisModel.labelFormatCode && text.trim() !== "" && Number.isFinite(number) ? formatChartNumber(number, categoryAxisModel.labelFormatCode) : text;

            tableSvg.push(element("text", [["class", "bar3d-table"], ["x", n(columnsLeft + (categorySlot(index) + 0.5) * column)], ["y", n(tableTop + tableRowPx / 2)], ["text-anchor", "middle"], ...textStyle], escapeText(label)));
        }

        for (const [row, series] of chart.series.entries()) {
            const y = tableTop + (row + 1.5) * tableRowPx;
            const nameX = tableLeft + 2 * LEGEND_TEXT_GAP;
            const valueFormat = table.valueFormatCode ?? (series.dataLabels ?? chart.dataLabels).valueFormat;

            if (table.keys) {
                tableSvg.push(
                    element("rect", [
                        ["x", n(nameX)],
                        ["y", n(y - keyPx / 2)],
                        ["width", n(keyPx)],
                        ["height", n(keyPx)],
                        ["fill", chart.colors[series.colorIndex % chart.colors.length]],
                    ]),
                );
            }

            tableSvg.push(element("text", [["class", "bar3d-table"], ["x", n(nameX + (table.keys ? keyPx + LEGEND_TEXT_GAP : 0))], ["y", n(y)], ...textStyle], escapeText(series.name)));

            for (const [index, value] of series.values.entries()) {
                tableSvg.push(element("text", [["class", "bar3d-table"], ["x", n(columnsLeft + (categorySlot(index) + 0.5) * column)], ["y", n(y)], ["text-anchor", "middle"], ...textStyle], escapeText(formatChartNumber(value, valueFormat))));
            }
        }

        // The horizontal rules between rows, the vertical ones between columns (the names' first), the outline.
        for (let row = 1; table.horizontalBorder && row <= chart.series.length; row++) {
            tableSvg.push(rule(tableLeft, tableTop + row * tableRowPx, tableRight, tableTop + row * tableRowPx));
        }

        for (let index = 0; table.verticalBorder && index < categoryCount; index++) {
            tableSvg.push(rule(columnsLeft + index * column, index === 0 ? tableTop + tableRowPx : tableTop, columnsLeft + index * column, tableBottom));
        }

        if (table.outline) {
            tableSvg.push(
                element("rect", [
                    ["class", "bar3d-table-rule"],
                    ["x", n(tableLeft)],
                    ["y", n(tableTop + tableRowPx)],
                    ["width", n(tableRight - tableLeft)],
                    ["height", n(tableBottom - tableTop - tableRowPx)],
                    ["fill", "none"],
                    ["stroke", table.borderColor],
                    ["stroke-width", n(pointsToPixels(table.borderWidth))],
                ]),
            );
        }
    }

    // The plot area's fill and border behind the box, its labels and the walls (normalizeChart keeps them only where the
    // target draws them, drawsPlotAreaOnEveryType, #565).
    const plotArea = chart.plotArea;
    const plotAreaBox =
        plotArea.fill.visible || plotArea.borderVisible
            ? [
                  element("rect", [
                      ["class", "bar3d-plot-area"],
                      ["x", n(area.x)],
                      ["y", n(area.y)],
                      ["width", n(area.width)],
                      ["height", n(area.height)],
                      ["fill", plotArea.fill.visible ? plotArea.fill.color : "none"],
                      ["stroke", plotArea.borderVisible ? plotArea.borderColor : undefined],
                      ["stroke-width", plotArea.borderVisible ? n(pointsToPixels(plotArea.borderWidth)) : undefined],
                  ]),
              ]
            : [];

    // titlePos places the title's top left corner; titleRotate turns it clockwise about its middle.
    const titleX = chart.titlePosition ? chart.titlePosition.x * width : width / 2;
    const titleY = chart.titlePosition ? chart.titlePosition.y * height + pointsToPixels(chart.titleFontSize) : pointsToPixels(chart.titleFontSize) + 4;
    const title = chart.showTitle
        ? [
              element(
                  "text",
                  [
                      ["class", "bar3d-title"],
                      ["x", n(titleX)],
                      ["y", n(titleY)],
                      ["text-anchor", chart.titlePosition ? "start" : "middle"],
                      ["transform", chart.titleRotate ? `rotate(${n(chart.titleRotate)} ${n(titleX)} ${n(titleY - pointsToPixels(chart.titleFontSize) / 2)})` : undefined],
                      ["font-family", chart.titleFontFace],
                      ["font-size", n(pointsToPixels(chart.titleFontSize))],
                      ["font-weight", chart.titleBold ? "bold" : "normal"],
                      ["fill", chart.titleColor],
                  ],
                  escapeText(chart.title),
              ),
          ]
        : [];

    return element(
        "svg",
        [
            ["xmlns", "http://www.w3.org/2000/svg"],
            ["viewBox", `0 0 ${n(width)} ${n(height)}`],
            ["data-chart-type", "bar3D"],
            ["data-bar3d-shape", chart.bar3DShape],
        ],
        element(
            "g",
            [
                ["font-family", chart.fontFace],
                ["font-size", n(pointsToPixels(chart.fontSize))],
                ["fill", "#000000"],
            ],
            [gradients, ...plotAreaBox, ...title, ...wall, floor, ...frame, ...grid, ...barSvg, ...crossingLabels, ...dataLabelSvg, ...categoryLabels, ...seriesLabels, ...tableSvg, ...legend],
        ),
    );
}
