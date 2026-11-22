import { PPTX_DEFAULTS } from "#src/defaults";
import { builtInFontMetrics } from "#src/fontMetrics";
import { automaticAxisInterval, type AxisScale, bar3DAxisScale } from "#src/chart/axisScale";
import { formatChartNumber } from "#src/chart/format";
import type { NormalizedChart } from "@pdflapas/model";
import { pointsToPixels } from "@pdflapas/model";
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
    const seriesCount = Math.max(1, chart.series.length);
    const stacked = chart.grouping === "stacked" || chart.grouping === "percentStacked";
    // The same text bands as the 2D plot area (chart.ts chartGrid): the legend and the category labels below the box.
    const gap = height * PPTX_DEFAULTS.chart.plotAreaGapRatio;
    const lineEm = PPTX_DEFAULTS.text.singleLineEm;
    const top = gap + (chart.showTitle ? pointsToPixels(chart.titleFontSize) * lineEm + gap : 0);
    const bottom = gap + (chart.showLegend ? pointsToPixels(chart.legendFontSize) * lineEm + gap : 0) + pointsToPixels(chart.fontSize) * lineEm;
    const side = width * PPTX_DEFAULTS.chart.bar3DSideMarginRatio;
    // Fits the box into the plot area. The depth is fixed (one series slot per category slot); the height is automatic, as
    // in LibreOffice (chart2 BarChart::getPreferredDiagramAspectRatio leaves Y free), so it is chosen to fill the area.
    const boxDepth = (stacked ? 1 : seriesCount) / categoryCount;
    const area = { x: side, y: top, width: width - 2 * side, height: height - top - bottom };
    const aspect = area.width / Math.max(1e-9, area.height);
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

    for (let index = 0; index <= tickCount; index++) {
        const value = Number((minimum + index * scale.interval).toPrecision(12));
        const label = project(sideX, frontDepth, value);
        const line = [label, project(sideX, backDepth, value), project(otherX, backDepth, value)];
        const labelX: Attribute[] =
            sideX === 0
                ? [
                      ["x", n(label.x - 8)],
                      ["text-anchor", "end"],
                  ]
                : [["x", n(label.x + 8)]];

        const tickLabel = axis ? formatChartNumber(value, axis.labelFormatCode) : String(value);

        grid.push(element("polyline", [["points", line.map(p).join(" ")], ...gridStroke]), element("text", [...labelX, ["y", n(label.y + labelFontSize / 4)], ...labelStyle], escapeText(tickLabel)));
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
        for (let series = seriesCount - 1; series >= 0; series--) {
            for (let category = 0; category < categoryCount; category++) {
                bars.push({ category, series, base: 0, value: chart.series[series]?.values[category] ?? 0 });
            }
        }
    }

    // gapWidth and gapDepth are percentages of the bar's width and depth (c:gapWidth, c:gapDepth); each category
    // has one slot along the category axis and, unless stacked, each series one slot along the depth axis.
    const halfWidth = 0.5 / (1 + chart.barGapWidthPercent / 100) / categoryCount;
    const depthSlot = stacked ? 1 : 1 / seriesCount;
    const depthThickness = depthSlot / (1 + chart.barGapDepthPercent / 100);
    const slotDepth = (series: number): number => (stacked ? 0.5 : (series + 0.5) * depthSlot);
    const barCenter = (bar: Bar): CameraPoint => project((bar.category + 0.5) / categoryCount, slotDepth(bar.series), (bar.base + bar.value) / 2);
    // Painter's order: the bar farthest from the camera first.
    const ordered: { bar: Bar; distance: number }[] = [];

    for (const bar of bars) {
        ordered.push({ bar, distance: barCenter(bar).distance });
    }

    ordered.sort((a, b) => b.distance - a.distance);

    const barSvg: Html[] = [];

    for (const { bar } of ordered) {
        const cx = (bar.category + 0.5) / categoryCount;
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
        const outline = series?.pointLineColors?.[bar.category];
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
            group.push(["stroke", outline], ["stroke-width", n(pointsToPixels(PPTX_DEFAULTS.chart.pointLineWidthPt))], ["stroke-linejoin", "round"]);
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

    const categoryLabels = Array.from({ length: categoryCount }, (_, index) => {
        const point = project((index + 0.5) / categoryCount, frontDepth, minimum);
        const label = labels[index] ?? String(index + 1);

        return element(
            "text",
            [
                ["x", n(point.x)],
                ["y", n(point.y + 20)],
                ["text-anchor", "middle"],
            ],
            escapeText(label),
        );
    });
    // LibreOffice 26.2.6 names the depth slots only when there are several: a single series or a stacked chart, whose
    // series share one slot, gets no depth-axis label. A target may name a single unstacked series (namesSingleDepthSlot).
    const namesSlots = !stacked && (chart.series.length > 1 || chart.namesSingleDepthSlot === true);
    const seriesLabels = (namesSlots ? chart.series : []).map((series, index) => {
        const point = project(otherX, slotDepth(index), minimum);

        return element(
            "text",
            [
                ["class", "bar3d-series-label"],
                ["x", n(point.x + 7)],
                ["y", n(point.y + 12)],
            ],
            escapeText(series.name),
        );
    });
    // The legend has its own font (legendFontFace/legendFontSize/legendColor), not the axis-label font of the <g>.
    const legendFont: Attribute[] = [
        ["font-family", chart.legendFontFace],
        ["font-size", n(pointsToPixels(chart.legendFontSize))],
        ["fill", chart.legendColor],
    ];
    const keySize = Math.round(pointsToPixels(chart.legendFontSize) * PPTX_DEFAULTS.chart.legendKeyFontRatio);
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

    const legend: Html[] = [];

    if (chart.showLegend) {
        for (const [index, { name, fill, stroke }] of legendEntries.entries()) {
            const itemWidth = Math.min(100, width / Math.max(2, legendEntries.length + 1));
            const x = width / 2 + (index - (legendEntries.length - 1) / 2) * itemWidth;
            const y = height - 12;

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

    const title = chart.showTitle
        ? [
              element(
                  "text",
                  [
                      ["class", "bar3d-title"],
                      ["x", n(width / 2)],
                      ["y", n(pointsToPixels(chart.titleFontSize) + 4)],
                      ["text-anchor", "middle"],
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
            [gradients, ...title, ...wall, floor, ...frame, ...grid, ...barSvg, ...categoryLabels, ...seriesLabels, ...legend],
        ),
    );
}
