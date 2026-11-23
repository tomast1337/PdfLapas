import { describe, expect, test } from "bun:test";
import { format } from "echarts/core";
import { inchesToPixels, pointsToPixels } from "@pdflapas/model";
import { JSDOM } from "jsdom";
import { renderedPage } from "./renderedPage";
import type PptxGenJS from "pptxgenjs";
import { UnsupportedChartError } from "../packages/pptxgenjs/src/chart/errors";
import { formatChartNumber } from "../packages/pptxgenjs/src/chart/format";
import { PPTX_DEFAULTS } from "../packages/pptxgenjs/src/defaults";
import { normalizeChart } from "../packages/pptxgenjs/src/normalize/chart";
import { normalizeColor } from "../packages/pptxgenjs/src/normalize/style";
import { PuppeteerGen, PuppeteerGenPowerPoint } from "../packages/pptxgenjs/src/node";
import { unresolvedModelHtml } from "../packages/pptxgenjs/src/render/document";

const categoricalData: PptxGenJS.OptsChartData[] = [
    { name: "First", labels: ["A", "B", "C"], values: [2, 5, 3] },
    { name: "Second", labels: ["A", "B", "C"], values: [4, 1, 6] },
];
// PptxGenJS scatter data: the x values first, then each y series, whose labels are its points' own labels.
const scatterData: PptxGenJS.OptsChartData[] = [
    { name: "X", values: [1, 2, 3] },
    { name: "Lorem", labels: ["Alpha", "Beta", "Gamma"], values: [4, 6, 5] },
];
const bubbleData: PptxGenJS.OptsChartData[] = [
    { name: "X", values: [1, 2, 3] },
    { name: "Lorem", values: [4, 6, 5], sizes: [11, 22, 33] },
];

describe("chart normalization and rendering", () => {
    test("normalizes a basic chart without backend-specific options", () => {
        const chart = normalizeChart("bar", categoricalData, {
            chartColors: ["4472C4", "ED7D31"],
            showLegend: true,
            legendPos: "b",
            barGrouping: "stacked",
        });

        expect(chart.type).toBe("bar");
        expect<string[]>(chart.colors).toEqual(["#4472C4", "#ED7D31"]);
        expect(chart.grouping).toBe("stacked");
        expect(chart.legendPosition).toBe("b");
        expect(chart.series.map(series => series.name)).toEqual(["First", "Second"]);
        expect(chart.valueAxisMinimum).toBe(0);
        expect(chart.valueAxisMaximum).toBe(10);
    });

    test("draws the title, axes and data labels in Arial when no font is set, as PptxGenJS writes them", () => {
        // PptxGenJS 4.0.1: `<a:latin typeface="${titleFontFace || 'Arial'}">`, and `|| 'Arial'` for axis and data labels.
        const chart = normalizeChart("bar", categoricalData, { showTitle: true, title: "T", showValAxisTitle: true, valAxisTitle: "V" });

        expect([chart.fontFace, chart.titleFontFace, chart.valueAxes[0]?.titleFontFace, chart.valueAxes[0]?.labelFontFace]).toEqual(["Arial", "Arial", "Arial", "Arial"]);
        expect(normalizeChart("bar", categoricalData, { titleFontFace: "Georgia" }).titleFontFace).toBe("Georgia");
    });

    test("ignores the chart-level fontFace, which PptxGenJS reads for no chart text", () => {
        // PptxGenJS 4.0.1 writes titleFontFace, valAxisTitleFontFace, valAxisLabelFontFace, catAxisLabelFontFace and
        // dataLabelFontFace `|| 'Arial'`, and legendFontFace only when set; `fontFace` reaches none of them.
        const chart = normalizeChart("bar", categoricalData, { fontFace: "Georgia", showTitle: true, title: "T", showValAxisTitle: true, valAxisTitle: "V" });

        expect([chart.fontFace, chart.titleFontFace, chart.valueAxes[0]?.titleFontFace, chart.valueAxes[0]?.labelFontFace, chart.legendFontFace]).toEqual(["Arial", "Arial", "Arial", "Arial", "Calibri"]);
        const axes = normalizeChart("bar", categoricalData, { fontFace: "Georgia", valAxisTitleFontFace: "Courier New", valAxisLabelFontFace: "Verdana" });

        expect([axes.valueAxes[0]?.titleFontFace, axes.valueAxes[0]?.labelFontFace]).toEqual(["Courier New", "Verdana"]);
    });

    test("ignores the chart-level fontSize, which PptxGenJS reads for no chart text", () => {
        // PptxGenJS 4.0.1 writes catAxisLabelFontSize, valAxisLabelFontSize and dataLabelFontSize `|| 12`, titleFontSize
        // `|| 18`, and valAxisTitleFontSize and legendFontSize only when set, leaving Office's 10pt: chart XML with
        // `fontSize: 31` is byte-identical to chart XML without.
        const chart = normalizeChart("bar", categoricalData, { fontSize: 31, showTitle: true, title: "T", showValAxisTitle: true, valAxisTitle: "V" });

        expect([chart.fontSize, chart.titleFontSize, chart.valueAxes[0]?.titleFontSize, chart.valueAxes[0]?.labelFontSize, chart.legendFontSize]).toEqual([12, 18, 10, 12, 10]);
        const axes = normalizeChart("bar", categoricalData, { fontSize: 31, valAxisLabelFontSize: 9, titleFontSize: 20 });

        expect([axes.valueAxes[0]?.labelFontSize, axes.titleFontSize]).toEqual([9, 20]);
    });

    test("legend ignores chart fontFace/fontSize like PptxGenJS and defaults to 10pt body font", () => {
        // PptxGenJS writes legend <c:txPr> only for legendFontFace/legendFontSize/legendColor.
        const inherited = normalizeChart("bar", categoricalData, { showLegend: true, fontFace: "Arial", fontSize: 20 });

        expect(inherited.legendFontFace).toBe("Calibri");
        expect(inherited.legendFontSize).toBe(10);

        const explicit = normalizeChart("bar", categoricalData, { showLegend: true, legendFontFace: "Georgia", legendFontSize: 14 });

        expect(explicit.legendFontFace).toBe("Georgia");
        expect(explicit.legendFontSize).toBe(14);
    });

    test("renders supported charts as inline SVG and remains chainable", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const result = slide.addChart("bar", categoricalData, {
            x: 1,
            y: 1,
            w: 5,
            h: 3,
            showLegend: true,
            objectName: "Sales chart",
            altText: "Quarterly sales",
        });

        expect(result).toBe(slide);
        const element = renderedPage(presentation).querySelector<HTMLElement>(".slide-chart");
        const svg = element?.querySelector("svg");

        expect(element?.dataset.objectName).toBe("Sales chart");
        expect(element?.getAttribute("aria-label")).toBe("Quarterly sales");
        expect(element?.style.left).toBe("96px");
        expect(element?.style.width).toBe("480px");
        expect(svg?.getAttribute("width")).toBe("100%");
        expect(svg?.querySelectorAll("path").length).toBeGreaterThan(0);
        expect(svg?.textContent).toContain("First");
    });

    test("keeps every axis label inside the chart frame, less the plot-area gap, as PowerPoint sizes the plot area", () => {
        // #443: "$120,000" on a column chart's value axis, long names on a horizontal bar chart's category axis and its last
        // value label, centered on the plot's right edge, overflowed the fixed 6% and 2% gutters and were clipped.
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const names = ["Lorem ipsum dolor", "Sit amet", "Consectetur elit"];
        const options = { y: 0.5, w: 4.4, h: 4.1, fontSize: 11, valAxisMinVal: 0, valAxisLabelFormatCode: "$#,##0" } as const;

        slide.addChart("bar", [{ name: "S", labels: names, values: [82400, 115200, 91200] }], { ...options, x: 0.4, barDir: "col", valAxisMaxVal: 120000, valAxisMajorUnit: 20000 });
        slide.addChart("bar", [{ name: "S", labels: names, values: [123600, 66100, 77400] }], { ...options, x: 5.2, barDir: "bar", valAxisMaxVal: 150000, valAxisMajorUnit: 50000 });
        const width = inchesToPixels(4.4);
        const gap = inchesToPixels(4.1) * PPTX_DEFAULTS.chart.plotAreaGapRatio;
        const extents: Record<string, { left: number; right: number }> = {};

        for (const svg of renderedPage(presentation).querySelectorAll(".slide-chart svg")) {
            for (const text of svg.querySelectorAll("text")) {
                const x = Number(/translate\(([-\d.]+)/.exec(text.getAttribute("transform") ?? "")?.[1]);
                // The width ECharts lays the label out with: zrender's text measure at the label's font size.
                const labelWidth = format.getTextRect(text.textContent ?? "", `${text.style.fontSize} ${text.style.fontFamily}`).width;
                const anchor = text.getAttribute("text-anchor");
                const left = anchor === "end" ? x - labelWidth : anchor === "middle" ? x - labelWidth / 2 : x;

                extents[text.textContent ?? ""] = { left, right: left + labelWidth };
            }
        }

        for (const label of ["$120,000", "Lorem ipsum dolor", "$150,000"]) {
            expect(extents[label]?.left).toBeGreaterThanOrEqual(gap - 1);
            expect(extents[label]?.right).toBeLessThanOrEqual(width - gap + 1);
        }

        // The widest label starts at the gap: the plot gives up no more room than the labels need.
        expect(extents["$120,000"]?.left).toBeLessThan(gap + 1);
        expect(extents["Lorem ipsum dolor"]?.left).toBeLessThan(gap + 1);
        expect(extents["$150,000"]?.right).toBeGreaterThan(width - gap - 1);
    });

    test("renders every initial 2D chart family", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("line", categoricalData, { x: 0, y: 0, w: 3, h: 2 });
        slide.addChart("area", categoricalData, { x: 3, y: 0, w: 3, h: 2 });
        slide.addChart("radar", categoricalData, { x: 6, y: 0, w: 3, h: 2 });
        slide.addChart("pie", [categoricalData[0]!], { x: 0, y: 2, w: 3, h: 2 });
        slide.addChart("doughnut", [categoricalData[0]!], { x: 3, y: 2, w: 3, h: 2 });
        const xyData: PptxGenJS.OptsChartData[] = [
            { name: "X", values: [1, 2, 3] },
            { name: "Y", values: [2, 5, 3], sizes: [2, 8, 4] },
        ];

        slide.addChart("scatter", xyData, { x: 6, y: 2, w: 3, h: 2 });
        slide.addChart("bubble", xyData, { x: 0, y: 4, w: 3, h: 1.5 });

        expect(renderedPage(presentation).querySelectorAll(".slide-chart")).toHaveLength(7);
        expect(renderedPage(presentation).querySelectorAll(".slide-chart svg")).toHaveLength(7);
    });

    test("sizes the pie in the plot area the title and legend leave and the hole as the target reads c:holeSize (#468)", () => {
        // The first slice's outer arc starts at 12 o'clock: its top is the ring's top, its x the ring's center.
        const ring = (presentation: PuppeteerGen, type: "pie" | "doughnut", options: PptxGenJS.IChartOpts) => {
            presentation.addSlide().addChart(type, [categoricalData[0]!], { x: 0, y: 0, w: 4.4, h: 4.1, ...options });
            const path = renderedPage(presentation).querySelector('.slide-chart path[ecmeta_ssr_type="chart"]');
            const numbers = (path?.getAttribute("d") ?? "").match(/-?[\d.]+/g)?.map(Number) ?? [];

            return { numbers, model: presentation.getModel().slides[0]?.objects[0] };
        };
        const width = inchesToPixels(4.4);
        const height = inchesToPixels(4.1);
        const layout = PPTX_DEFAULTS.chart.pieLayout;
        const inset = pointsToPixels(layout.insetPt);
        // A bottom legend: two insets and 1.05 legend font sizes below the ring, an inset above it.
        const radius = (height - inset - (2 * inset + pointsToPixels(10) * layout.legendBandEm)) / 2;
        const pie = ring(new PuppeteerGen(), "pie", { showLegend: true, legendPos: "b" }).numbers;

        // ECharts starts the slice at the arc's top (M cx top), then its radius (A r r).
        expect(pie[0]).toBeCloseTo(width / 2, 1);
        expect(pie[1]).toBeCloseTo(inset, 1);
        expect(pie[2]).toBeCloseTo(radius, 1);

        // A title band: 2% of the height, 1.7 title font sizes and 14pt.
        const titled = ring(new PuppeteerGen(), "pie", { showTitle: true, title: "T" }).numbers;
        const titleBand = height * PPTX_DEFAULTS.chart.plotAreaGapRatio + pointsToPixels(18) * layout.titleBandEm + pointsToPixels(layout.titleBandPt);

        expect(titled[1]).toBeCloseTo(titleBand, 1);
        expect(titled[2]).toBeCloseTo((height - titleBand - inset) / 2, 1);

        // LibreOffice draws every hole at half the ring; PowerPoint reads holeSize.
        const holeSize = (presentation: PuppeteerGen) => {
            const object = ring(presentation, "doughnut", { holeSize: 75 }).model;

            return object?.kind === "chart" ? object.chart.doughnutHoleSize : undefined;
        };

        expect(holeSize(new PuppeteerGen())).toBe(PPTX_DEFAULTS.chart.doughnutHoleSizePct);
        expect(holeSize(new PuppeteerGenPowerPoint())).toBe(75);
    });

    test("draws no slice labels on pie and doughnut charts, as PptxGenJS's per-point c:dLbl turn them off", () => {
        // PptxGenJS 4.0.1 writes, for each point, `<c:dLbl>` with showVal, showCatName and showPercent "0" when
        // showValue, showLabel and showPercent are unset: PowerPoint and LibreOffice draw no slice labels.
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("pie", [categoricalData[0]!], { x: 0, y: 0, w: 3, h: 2 });
        slide.addChart("doughnut", [categoricalData[0]!], { x: 3, y: 0, w: 3, h: 2, showLegend: true });
        const [pie, doughnut] = renderedPage(presentation).querySelectorAll(".slide-chart svg");

        expect(pie?.querySelectorAll("text")).toHaveLength(0);
        // Only the legend names the categories.
        expect(Array.from(doughnut?.querySelectorAll("text") ?? [], text => text.textContent)).toEqual(["A", "B", "C"]);
    });

    test("reports unsupported types and options with structured details", () => {
        try {
            normalizeChart("bar", categoricalData, { dataNoEffects: true } as unknown as PptxGenJS.IChartOpts);
            throw new Error("Expected normalization to fail");
        } catch (error) {
            expect(error).toBeInstanceOf(UnsupportedChartError);
            expect((error as UnsupportedChartError).code).toBe("UNSUPPORTED_CHART");
            expect((error as UnsupportedChartError).reason).toBe("chart-option");
            expect((error as UnsupportedChartError).unsupportedOptions).toEqual(["dataNoEffects"]);
        }
    });

    test("normalizes data labels from PptxGenJS's defaults and the positions its option cleanup keeps", () => {
        // PptxGenJS 4.0.1: dataLabelFontSize || 12, dataLabelColor || '000000', dataLabelFontFace || 'Arial', format
        // '#,##0' (pie and doughnut: showPercent ? '0%' : 'General'), no c:dLblPos unless dataLabelPosition is kept.
        const bar = normalizeChart("bar", categoricalData);

        expect(bar.dataLabels).toEqual({
            showValue: false,
            showCategory: false,
            showPercent: false,
            position: "outEnd",
            valueFormat: "#,##0",
            percentFormat: "#,##0",
            color: "#000000" as typeof bar.dataLabels.color,
            fontFace: "Arial",
            fontSize: 12,
            bold: false,
            italic: false,
        });
        expect(normalizeChart("bar", categoricalData, { barGrouping: "stacked" }).dataLabels.position).toBe("ctr");
        expect(normalizeChart("line", categoricalData).dataLabels.position).toBe("r");
        expect(normalizeChart("pie", [categoricalData[0]!], {}, "powerpoint").dataLabels.position).toBe("bestFit");
        // PptxGenJS deletes outEnd on a clustered bar, any position on a doughnut, and a bar position on a line chart.
        expect(normalizeChart("bar", categoricalData, { showValue: true, dataLabelPosition: "outEnd" }).dataLabels.position).toBe("outEnd");
        expect(normalizeChart("bar", categoricalData, { showValue: true, dataLabelPosition: "inEnd" }).dataLabels.position).toBe("inEnd");
        expect(normalizeChart("bar", categoricalData, { showValue: true, barGrouping: "stacked", dataLabelPosition: "outEnd" }).dataLabels.position).toBe("outEnd");
        expect(normalizeChart("doughnut", [categoricalData[0]!], { showValue: true, dataLabelPosition: "t" }).dataLabels.position).toBe("ctr");
        expect(normalizeChart("line", categoricalData, { showValue: true, dataLabelPosition: "inEnd" }).dataLabels.position).toBe("r");
        expect(normalizeChart("line", categoricalData, { showValue: true, dataLabelPosition: "t" }).dataLabels.position).toBe("t");

        const styled = normalizeChart("line", categoricalData, { showValue: true, dataLabelColor: "C00000", dataLabelFontFace: "Courier New", dataLabelFontSize: 9, dataLabelFontBold: true, dataLabelFontItalic: true, dataLabelFormatCode: "0.0%" });

        expect([styled.dataLabels.color, styled.dataLabels.fontFace, styled.dataLabels.fontSize, styled.dataLabels.bold, styled.dataLabels.italic, styled.dataLabels.valueFormat]).toEqual([
            "#C00000" as typeof styled.dataLabels.color,
            "Courier New",
            9,
            true,
            true,
            "0.0%",
        ]);
        // showLabel and showPercent are written only on pie and doughnut points.
        expect(normalizeChart("bar", categoricalData, { showLabel: true, showPercent: true }).dataLabels).toMatchObject({ showCategory: false, showPercent: false });

        // A pie formats the percentage with the code and the value in General once showPercent is set (LibreOffice).
        const pie = normalizeChart("pie", [categoricalData[0]!], { showLabel: true, showValue: true, showPercent: true });

        expect(pie.dataLabels).toMatchObject({ showCategory: true, showValue: true, showPercent: true, valueFormat: "General", percentFormat: "0%" });
        expect(normalizeChart("pie", [categoricalData[0]!], { showValue: true }).dataLabels).toMatchObject({ valueFormat: "General" });
    });

    test("normalizes data labels on every chart type, with the positions PptxGenJS keeps (#477)", () => {
        // PptxGenJS deletes every dataLabelPosition on area, radar, bar3D and doughnut charts; LibreOffice centers an
        // area label, puts a radar one right of its point and a bar3D one above its bar.
        expect(normalizeChart("area", categoricalData, { showValue: true, dataLabelPosition: "t" }).dataLabels).toMatchObject({ showValue: true, position: "ctr" });
        expect(normalizeChart("radar", categoricalData, { showValue: true, dataLabelPosition: "t" }).dataLabels).toMatchObject({ showValue: true, position: "r" });
        expect(normalizeChart("bar3D", categoricalData, { showValue: true, dataLabelPosition: "ctr" }).dataLabels).toMatchObject({ showValue: true, position: "outEnd" });
        // A scatter point's label: right of it, its y value in General, or its own label for showLabel.
        expect(normalizeChart("scatter", scatterData, { showValue: true }).dataLabels).toMatchObject({ position: "r", valueFormat: "General" });
        expect(normalizeChart("scatter", scatterData, { showLabel: true, dataLabelPosition: "t" }).dataLabels).toMatchObject({ showCategory: true, position: "t" });
        expect(normalizeChart("bubble", bubbleData, { showLabel: true }).dataLabels.showCategory).toBe(false);
        // showSerName on every chart, a pie's too: PptxGenJS writes it on each point's c:dLbl (#565).
        expect(normalizeChart("bar", categoricalData, { showSerName: true }).dataLabels.showSeriesName).toBe(true);
        expect(normalizeChart("pie", [categoricalData[0]!], { showSerName: true }).dataLabels.showSeriesName).toBe(true);
        // showLeaderLines and showLegendKey are accepted: Office draws no leader line to an unmoved label, and
        // PptxGenJS always writes showLegendKey 0.
        expect(() => normalizeChart("pie", [categoricalData[0]!], { showValue: true, showLeaderLines: true, showLegendKey: true } as PptxGenJS.IChartOpts)).not.toThrow();
        expect(() => normalizeChart("bar", categoricalData, { showValue: true, dataLabelFormatCode: "0;[Red]-0" })).toThrow("dataLabelFormatCode[format]");
    });

    test("gives each combo chart type its own data labels (#477)", () => {
        const types = [
            { type: "bar", data: categoricalData, options: { showValue: true, dataLabelColor: "C00000" } },
            { type: "line", data: categoricalData, options: { secondaryValAxis: true, showSerName: true, dataLabelPosition: "t" } },
        ] as PptxGenJS.IChartMulti[];
        const chart = normalizeChart(types, [], { valAxes: [{}, {}], showValue: false } as PptxGenJS.IChartOpts);
        const bar = chart.series.find(series => series.type === "bar");
        const line = chart.series.find(series => series.type === "line");

        expect(bar?.dataLabels).toMatchObject({ showValue: true, color: "#C00000", position: "outEnd" });
        expect(line?.dataLabels).toMatchObject({ showValue: false, showSeriesName: true, position: "t" });
    });

    test("draws area, radar, scatter, bubble and bar3D data labels with the series name and the value", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const one = [{ name: "North", labels: ["A", "B", "C"], values: [3, 5, 4] }];

        slide.addChart("area", one, { x: 0, y: 0, w: 3, h: 2, showValue: true, showSerName: true });
        slide.addChart("radar", one, { x: 3, y: 0, w: 3, h: 2, showValue: true });
        slide.addChart("scatter", scatterData, { x: 6, y: 0, w: 3, h: 2, showLabel: true });
        slide.addChart("bubble", bubbleData, { x: 0, y: 2, w: 3, h: 2, showValue: true });
        slide.addChart("bar3D", one, { x: 3, y: 2, w: 3, h: 2, showValue: true });
        const charts = Array.from(renderedPage(presentation).querySelectorAll(".slide-chart svg"));
        const texts = (index: number) => Array.from(charts[index]?.querySelectorAll("text") ?? [], text => text.textContent ?? "");

        expect(texts(0)).toContain("North; 5");
        expect(texts(1)).toEqual(expect.arrayContaining(["3", "5", "4"]));
        expect(texts(2)).toEqual(expect.arrayContaining(["Alpha", "Beta", "Gamma"]));
        // LibreOffice prints a bubble's size for showValue (bubbleValueLabelShowsSize).
        expect(texts(3)).toEqual(expect.arrayContaining((bubbleData[1]!.sizes ?? []).map(String)));
        expect(charts[4]?.querySelectorAll(".bar3d-data-label")).toHaveLength(3);
    });

    test("joins a series name and a value with the target's separator", () => {
        const one = [{ name: "North", labels: ["A", "B", "C"], values: [3, 5, 4] }];
        const label = (presentation: PuppeteerGen) => {
            presentation.addSlide().addChart("bar", one, { x: 0, y: 0, w: 3, h: 2, showValue: true, showSerName: true });

            return Array.from(renderedPage(presentation).querySelectorAll(".slide-chart svg text"), text => text.textContent ?? "");
        };

        // LibreOffice prints "North; 5", PowerPoint "North, 5" (#477, #568).
        expect(label(new PuppeteerGen())).toContain("North; 5");
        expect(label(new PuppeteerGenPowerPoint())).toContain("North, 5");
    });

    test("draws each bar and line value and each slice's category and percentage in the data-label style", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const thousands = [{ name: "S", labels: ["A", "B"], values: [1200, 3400] }];

        slide.addChart("bar", thousands, { x: 0, y: 0, w: 3, h: 2, showValue: true, dataLabelFormatCode: "$#,##0", dataLabelColor: "C00000", dataLabelFontSize: 9 });
        slide.addChart("bar", thousands, { x: 3, y: 0, w: 3, h: 2, barGrouping: "percentStacked", showValue: true });
        slide.addChart("line", thousands, { x: 6, y: 0, w: 3, h: 2, showValue: true, dataLabelPosition: "t" });
        slide.addChart("pie", [{ name: "S", labels: ["Lorem", "Ipsum"], values: [1, 3] }], { x: 0, y: 2, w: 3, h: 2, showLabel: true, showPercent: true });
        const charts = Array.from(renderedPage(presentation).querySelectorAll(".slide-chart svg"));
        const texts = (index: number) => Array.from(charts[index]?.querySelectorAll("text") ?? [], text => text.textContent ?? "");
        const labelText = (index: number, content: string) => Array.from(charts[index]?.querySelectorAll("text") ?? []).find(text => text.textContent === content);

        expect(texts(0)).toEqual(expect.arrayContaining(["$1,200", "$3,400"]));
        expect(labelText(0, "$1,200")?.getAttribute("style")).toContain("12px");
        expect(labelText(0, "$1,200")?.getAttribute("fill")).toBe("#C00000");
        // A percent-stacked bar draws its share but labels its own value.
        expect(texts(1)).toEqual(expect.arrayContaining(["1,200", "3,400"]));
        expect(texts(2)).toEqual(expect.arrayContaining(["1,200", "3,400"]));
        // Each field on its own line: ECharts writes one text element per line.
        expect(texts(3)).toEqual(["Lorem", "25%", "Ipsum", "75%"]);
    });

    test("places a pie's inEnd and outEnd labels at the rim as LibreOffice 26.2 does (#478)", () => {
        // A 4.4 x 4.1in pie (values 40, 25, 20, 15, 12pt values): LibreOffice draws the outEnd ring 332.5px wide at
        // (213.2, 192.5) in the chart, the "20" label (bisector at 9 o'clock) left of a point 1.5mm outside the rim,
        // and keeps the inEnd ring 367px wide at (211.2, 196.8), "20" right of a point 1.5mm inside it.
        expect(normalizeChart("pie", [categoricalData[0]!], { showValue: true, dataLabelPosition: "outEnd" }).dataLabels.position).toBe("outEnd");
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const data = [{ name: "S", labels: ["Lorem", "Ipsum", "Dolor", "Sit"], values: [40, 25, 20, 15] }];

        slide.addChart("pie", data, { x: 0, y: 0, w: 4.4, h: 4.1, showValue: true, dataLabelPosition: "outEnd" });
        slide.addChart("pie", data, { x: 5, y: 0, w: 4.4, h: 4.1, showValue: true, dataLabelPosition: "inEnd" });
        const charts = Array.from(renderedPage(presentation).querySelectorAll(".slide-chart svg"));
        const label = (index: number, content: string) => {
            const text = Array.from(charts[index]?.querySelectorAll("text") ?? []).find(candidate => candidate.textContent === content);
            const [x = Number.NaN, y = Number.NaN] = (/translate\(([-\d.]+) ([-\d.]+)\)/.exec(text?.getAttribute("transform") ?? "") ?? []).slice(1).map(Number);

            return { anchor: text?.getAttribute("text-anchor"), x, y };
        };
        const offset = pointsToPixels(PPTX_DEFAULTS.chart.pieLabel.offsetPt);
        const outside = label(0, "20");
        const inside = label(1, "20");

        expect(outside.anchor).toBe("end");
        expect(Math.abs(outside.x - (213.2 - 166.2 - offset))).toBeLessThan(1.5);
        expect(Math.abs(outside.y - 192.5)).toBeLessThan(1.5);
        // The 40 slice's bisector points up-right: its outEnd label starts right of its anchor, its inEnd label ends there.
        expect(label(0, "40").anchor).toBe("start");
        expect(label(1, "40").anchor).toBe("end");
        expect(inside.anchor).toBe("start");
        expect(Math.abs(inside.x - (211.2 - 183.5 + offset))).toBeLessThan(1);
        expect(Math.abs(inside.y - 196.8)).toBeLessThan(1);
    });

    test("places a pie's bestFit labels near the rim, or outside a slice too narrow for them (#558)", () => {
        // A 4.4 x 4.1in pie without title or legend: a 183.8px ring at (211.2, 196.8). LibreOffice centers explicit
        // bestFit labels 167 to 170px from its center, PowerPoint (with no dataLabelPosition) 162 to 168px.
        const data = [{ name: "S", labels: ["Lorem", "Ipsum", "Dolor", "Sit"], values: [40, 25, 20, 15] }];
        const labels = (presentation: PuppeteerGen, content: string) => {
            const found: { anchor: string | null; distance: number }[] = [];

            for (const text of renderedPage(presentation).querySelectorAll(".slide-chart svg text")) {
                const [x = Number.NaN, y = Number.NaN] = (/translate\(([-\d.]+) ([-\d.]+)\)/.exec(text.getAttribute("transform") ?? "") ?? []).slice(1).map(Number);

                if (text.textContent === content) {
                    found.push({ anchor: text.getAttribute("text-anchor"), distance: Math.hypot(x - 211.2, y - 196.8) });
                }
            }

            return found;
        };
        const libreOffice = new PuppeteerGen();
        const powerPoint = new PuppeteerGenPowerPoint();

        libreOffice.addSlide().addChart("pie", data, { x: 0, y: 0, w: 4.4, h: 4.1, showValue: true, dataLabelPosition: "bestFit" });
        powerPoint.addSlide().addChart("pie", data, { x: 0, y: 0, w: 4.4, h: 4.1, showValue: true });

        for (const presentation of [libreOffice, powerPoint]) {
            for (const value of ["40", "25", "20", "15"]) {
                const [label] = labels(presentation, value);

                expect(label?.anchor).toBe("middle");
                expect(label?.distance).toBeGreaterThan(163);
                expect(label?.distance).toBeLessThan(171);
            }
        }

        // With no dataLabelPosition LibreOffice draws them at ctr's mid-radius (pieUnsetLabelPosition).
        const unset = new PuppeteerGen();

        unset.addSlide().addChart("pie", data, { x: 0, y: 0, w: 4.4, h: 4.1, showValue: true });
        expect(unset.getModel().slides[0]?.objects[0]).toMatchObject({ chart: { dataLabels: { position: "ctr" } } });
        expect(powerPoint.getModel().slides[0]?.objects[0]).toMatchObject({ chart: { dataLabels: { position: "bestFit" } } });
        expect(normalizeChart("pie", data, { showValue: true }, "libreoffice").dataLabels.position).toBe("ctr");
        expect(normalizeChart("pie", data, { showValue: true, dataLabelPosition: "bestFit" }, "libreoffice").dataLabels.position).toBe("bestFit");

        // A 0.72 degree slice is too narrow for its label, which goes outside the ring, shrunk to make room for it (#572).
        const narrow = new PuppeteerGen();

        narrow.addSlide().addChart("pie", [{ name: "S", labels: ["Lorem", "Ipsum"], values: [499, 1] }], { x: 0, y: 0, w: 4.4, h: 4.1, showValue: true, dataLabelPosition: "bestFit" });
        const sector = renderedPage(narrow).querySelector(".slide-chart svg path")?.getAttribute("d") ?? "";
        const [radius = Number.NaN] = (/A([\d.]+) /.exec(sector) ?? []).slice(1).map(Number);
        const [centerX = Number.NaN, centerY = Number.NaN] = (/L([\d.]+) ([\d.]+)Z/.exec(sector) ?? []).slice(1).map(Number);
        const fromCenter = (content: string) => {
            const text = Array.from(renderedPage(narrow).querySelectorAll(".slide-chart svg text")).find(candidate => candidate.textContent === content);
            const [x = Number.NaN, y = Number.NaN] = (/translate\(([-\d.]+) ([-\d.]+)\)/.exec(text?.getAttribute("transform") ?? "") ?? []).slice(1).map(Number);

            return Math.hypot(x - centerX, y - centerY);
        };

        expect(radius).toBeLessThan(183.8);
        expect(fromCenter("499")).toBeLessThan(radius);
        expect(fromCenter("1")).toBeGreaterThan(radius);
    });

    test("shrinks a bestFit pie's ring for each label too wide for its slice of LibreOffice's first, small ring, and keeps it for PowerPoint (#572)", () => {
        // LibreOffice 26.2.7 on a 6 x 4.4in pie, 12pt Arial percentages at 96 dpi: a 40/25/20/15 ring 395px wide, a
        // 70/15/8/4/3 one 373px, though every label of both lies in its slice. PowerPoint keeps both at 393px.
        const ringWidth = (values: number[], presentation: PuppeteerGen = new PuppeteerGen()) => {
            presentation.addSlide().addChart("pie", [{ name: "S", labels: values.map(String), values }], { x: 0, y: 0, w: 6, h: 4.4, showPercent: true, showValue: false, dataLabelPosition: "bestFit", dataLabelFontFace: "Arial" });
            let radius = 0;

            for (const path of renderedPage(presentation).querySelectorAll(".slide-chart svg path")) {
                radius = Math.max(radius, Number(/A([\d.]+) /.exec(path.getAttribute("d") ?? "")?.[1] ?? 0));
            }

            return 2 * radius;
        };

        expect(ringWidth([40, 25, 20, 15])).toBeCloseTo(395, -1);
        expect(ringWidth([70, 15, 8, 4, 3])).toBeCloseTo(373, -1);
        expect(ringWidth([70, 15, 8, 4, 3], new PuppeteerGenPowerPoint())).toBeCloseTo(395, -1);
    });

    test("wraps a long outEnd pie label where LibreOffice does, and keeps one line per field for PowerPoint (#558)", () => {
        const data = [{ name: "S", labels: ["Lorem", "Ipsum", "Dolor", "Sit"], values: [40, 25, 20, 15] }];
        const texts = (presentation: PuppeteerGen) => {
            presentation.addSlide().addChart("pie", data, { x: 0, y: 0, w: 4.4, h: 4.1, showLabel: true, showPercent: true, showValue: false, dataLabelPosition: "outEnd", dataLabelFontFace: "Arial" });

            return Array.from(renderedPage(presentation).querySelectorAll(".slide-chart svg text"), text => text.textContent);
        };

        // LibreOffice 26.2.6 wraps the label on the right to 80% of the room beside its rim point; the left ones fit.
        expect(texts(new PuppeteerGen())).toEqual(expect.arrayContaining(["Lor", "em", "40", "%", "Dolor", "20%"]));
        expect(texts(new PuppeteerGenPowerPoint())).toEqual(expect.arrayContaining(["Lorem", "40%", "Dolor", "20%"]));
    });

    test("draws a combo chart's areas behind its bars, on the ticks, and lists them where the target does (#538)", () => {
        const labels = ["Q1", "Q2", "Q3", "Q4"];
        const bars: PptxGenJS.IChartMulti = { type: "bar", data: [{ name: "A", labels, values: [3, 5, 4, 7] }], options: {} };
        const areas: PptxGenJS.IChartMulti = { type: "area", data: [{ name: "R", labels, values: [2, 3, 2, 3] }], options: {} };
        const drawn = (target: "libreoffice" | "powerpoint", types: PptxGenJS.IChartMulti[]) => {
            const presentation = new PuppeteerGen(undefined, { target });

            // `as Function`: PptxGenJS types the second argument as data; a combo call passes the options there.
            (presentation.addSlide().addChart as Function)(types, { x: 0, y: 0, w: 4.4, h: 4.1, chartColors: ["4472C4", "ED7D31"], showLegend: true, legendPos: "b" });
            const object = presentation.getModel().slides[0]?.objects[0];
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");
            // Each type in paint order: a bar is a four-sided path, an area one with a side per point and its base.
            const painted: string[] = [];

            for (const path of svg?.querySelectorAll("path") ?? []) {
                const kind = (path.getAttribute("d")?.match(/L/g) ?? []).length > 4 ? "area" : "bar";
                const fill = path.getAttribute("fill");

                if ((fill === "#4472C4" || fill === "#ED7D31") && !painted.includes(kind)) {
                    painted.push(kind);
                }
            }

            const legend: string[] = [];

            for (const text of svg?.querySelectorAll("text") ?? []) {
                if (text.textContent === "A" || text.textContent === "R") {
                    legend.push(text.textContent);
                }
            }

            return { painted, legend, chart: object?.kind === "chart" ? { categoriesOnTicks: object.chart.categoriesOnTicks, legendListsAreasFirst: object.chart.legendListsAreasFirst } : undefined };
        };

        expect(drawn("powerpoint", [bars, areas])).toEqual({ painted: ["area", "bar"], legend: ["A", "R"], chart: { categoriesOnTicks: true, legendListsAreasFirst: undefined } });
        // LibreOffice lists the areas first and keeps a bar-first combo's categories between the ticks.
        expect(drawn("libreoffice", [bars, areas])).toEqual({ painted: ["area", "bar"], legend: ["R", "A"], chart: { categoriesOnTicks: undefined, legendListsAreasFirst: true } });
        expect(drawn("libreoffice", [areas, bars]).chart).toEqual({ categoriesOnTicks: true, legendListsAreasFirst: true });
        // An unstacked area group's series overlap: LibreOffice draws the first in front (firstAreaSeriesInFront).
        const overlapping = (target: "libreoffice" | "powerpoint") => {
            const presentation = new PuppeteerGen(undefined, { target });
            const twoAreas = { ...areas, data: [...areas.data, { name: "S", labels, values: [1, 2, 1, 2] }] };

            (presentation.addSlide().addChart as Function)([{ type: "line", data: [{ name: "L", labels, values: [4, 6, 5, 7] }], options: {} }, twoAreas], { x: 0, y: 0, w: 4.4, h: 4.1, chartColors: ["4472C4", "ED7D31"] });
            const fills: string[] = [];

            for (const path of renderedPage(presentation).querySelectorAll(".slide-chart path")) {
                const fill = path.getAttribute("fill");

                if ((fill === "#4472C4" || fill === "#ED7D31") && (path.getAttribute("d")?.match(/L/g) ?? []).length > 4) {
                    fills.push(fill);
                }
            }

            return fills;
        };

        expect(overlapping("libreoffice")).toEqual(["#ED7D31", "#4472C4"]);
        expect(overlapping("powerpoint")).toEqual(["#4472C4", "#ED7D31"]);
        // PptxGenJS writes midCat only on a combo with an area type.
        expect(normalizeChart([bars, areas], []).categoriesOnTicks).toBe(true);
        expect(normalizeChart([bars, { ...areas, type: "line" }], []).categoriesOnTicks).toBeUndefined();
        expect(normalizeChart("area", categoricalData).categoriesOnTicks).toBeUndefined();
    });

    test("puts a negative bar's label past its own end, as LibreOffice 26.2 does (#478)", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const data = [{ name: "S", labels: ["A", "B", "C", "D"], values: [5, -3, 7, -6] }];

        slide.addChart("bar", data, { x: 0, y: 0, w: 4.4, h: 4.1, showValue: true });
        slide.addChart("bar", data, { x: 5, y: 0, w: 4.4, h: 4.1, showValue: true, barDir: "bar" });
        slide.addChart("bar", data, { x: 0, y: 4, w: 4.4, h: 4.1, showValue: true, dataLabelPosition: "inEnd" });
        const charts = Array.from(renderedPage(presentation).querySelectorAll(".slide-chart svg"));
        // Each text's anchor (x, y): the axis labels -8..8 by 2 come first, the odd data labels are unique.
        const position = (index: number, content: string) => {
            const text = Array.from(charts[index]?.querySelectorAll("text") ?? []).find(candidate => candidate.textContent === content);
            const [x = Number.NaN, y = Number.NaN] = (/translate\(([-\d.]+) ([-\d.]+)\)/.exec(text?.getAttribute("transform") ?? "") ?? []).slice(1).map(Number);

            return { x, y };
        };
        const zero = position(0, "0");
        const unit = (zero.y - position(0, "2").y) / 2;

        // outEnd (the clustered default): below the -3 column's end, above the 5 column's.
        expect(position(0, "-3").y).toBeGreaterThan(zero.y + 3 * unit);
        expect(position(0, "5").y).toBeLessThan(zero.y - 5 * unit);
        // A horizontal bar's: left of the -3 bar's end (the value axis is the second set of axis labels).
        const barZero = Array.from(charts[1]?.querySelectorAll("text") ?? []).find(text => text.textContent === "0" && text.getAttribute("text-anchor") === "middle");
        const barZeroX = Number(/translate\(([-\d.]+)/.exec(barZero?.getAttribute("transform") ?? "")?.[1]);

        expect(position(1, "-3").x).toBeLessThan(barZeroX - 3 * unit);
        // inEnd: inside the -3 column, at its end.
        expect(position(2, "-3").y).toBeGreaterThan(zero.y + 1.5 * unit);
        expect(position(2, "-3").y).toBeLessThan(zero.y + 3 * unit);
    });

    test("normalizes the category axis from PptxGenJS's makeCatAxis defaults", () => {
        // PptxGenJS 4.0.1: catAxisLabelFontSize || 12, catAxisLabelColor || '000000', catAxisLabelFontFace || 'Arial',
        // catAxisTitle || 'Axis Title' with no sz (Office's 10pt), and catGridLine { style: 'none' } unless set.
        const chart = normalizeChart("bar", categoricalData);

        expect(chart.categoryAxis).toEqual({
            showTitle: false,
            title: "Axis Title",
            titleColor: "#000000" as typeof chart.categoryAxis.titleColor,
            titleFontFace: "Arial",
            titleFontSize: 10,
            titleRotate: undefined,
            labelColor: "#000000" as typeof chart.categoryAxis.labelColor,
            labelFontFace: "Arial",
            labelFontSize: 12,
            labelBold: false,
            labelItalic: false,
            hidden: false,
            labelPosition: "low",
            labelRotate: 0,
            lineColor: "#888888" as typeof chart.categoryAxis.lineColor,
            lineWidth: 1,
            lineStyle: "solid",
            minorTickMark: "none",
            gridLineColor: "#888888" as typeof chart.categoryAxis.gridLineColor,
            gridLineWidth: 1,
            gridLineStyle: "none",
            gridLineCap: "flat",
        });

        const styled = normalizeChart("line", categoricalData, {
            showCatAxisTitle: true,
            catAxisTitle: "Quarter",
            catAxisTitleColor: "17365D",
            catAxisTitleFontFace: "Courier New",
            catAxisTitleFontSize: 14,
            catAxisTitleRotate: 10,
            catAxisLabelColor: "C00000",
            catAxisLabelFontFace: "Verdana",
            catAxisLabelFontSize: 9,
            catAxisLabelFontBold: true,
            catAxisLabelFontItalic: true,
            catGridLine: { color: "A5A5A5", size: 2, style: "dash", cap: "round" },
        });

        expect(styled.categoryAxis).toMatchObject({
            showTitle: true,
            title: "Quarter",
            titleFontFace: "Courier New",
            titleFontSize: 14,
            titleRotate: 10,
            labelFontFace: "Verdana",
            labelFontSize: 9,
            labelBold: true,
            labelItalic: true,
            gridLineWidth: 2,
            gridLineStyle: "dash",
            gridLineCap: "round",
        });
        expect([styled.categoryAxis.titleColor, styled.categoryAxis.labelColor, styled.categoryAxis.gridLineColor]).toEqual(["#17365D", "#C00000", "#A5A5A5"] as (typeof styled.categoryAxis.titleColor)[]);
        // A catGridLine without a style is solid (createGridLineElement: `style || 'solid'`).
        expect(normalizeChart("bar", categoricalData, { catGridLine: { color: "FF0000" } }).categoryAxis.gridLineStyle).toBe("solid");
        // A pie or doughnut has no axes in PptxGenJS's output: the options draw nothing there.
        expect(normalizeChart("pie", [categoricalData[0]!], { catAxisLabelColor: "FF0000" }).categoryAxis.labelColor).toBe("#FF0000" as typeof chart.categoryAxis.labelColor);
    });

    test("rejects category-axis options on chart types whose category axis it does not draw", () => {
        expect(() => normalizeChart("radar", categoricalData, { catGridLine: { color: "FF0000" } })).toThrow("catGridLine[radar]");
        // LibreOffice draws no radar category title: accepted, as it draws nothing.
        expect(() => normalizeChart("radar", categoricalData, { showCatAxisTitle: true, catAxisTitle: "Lorem" })).not.toThrow();
        expect(() => normalizeChart("bar", categoricalData, { catGridLine: { color: "FF0000", width: 2 } as PptxGenJS.OptsChartGridLine })).toThrow("catGridLine.width");
        expect(() => normalizeChart("bar", categoricalData, { catAxisLabelFontSize: -1 })).toThrow("catAxisLabelFontSize");
    });

    test("draws the category labels, title and gridlines in their own style and lowers the plot for the title", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const options = { y: 0, w: 4, h: 3 } as const;

        slide.addChart("bar", categoricalData, { ...options, x: 0 });
        slide.addChart("bar", categoricalData, {
            ...options,
            x: 4,
            catAxisLabelColor: "C00000",
            catAxisLabelFontSize: 14,
            showCatAxisTitle: true,
            catAxisTitle: "Quarter",
            catAxisTitleColor: "17365D",
            catGridLine: { color: "A5A5A5", style: "dash" },
        });
        const [plain, styled] = Array.from(renderedPage(presentation).querySelectorAll(".slide-chart svg"));
        const text = (svg: Element | undefined, content: string) => Array.from(svg?.querySelectorAll("text") ?? []).find(element => element.textContent === content);

        expect(text(styled, "A")?.getAttribute("fill")).toBe("#C00000");
        expect(text(styled, "A")?.getAttribute("style")).toContain(`${pointsToPixels(14)}px`);
        expect(text(styled, "Quarter")?.getAttribute("fill")).toBe("#17365D");
        expect(text(plain, "Axis Title")).toBeUndefined();
        // The category labels sit higher, above the title's band.
        expect(
            Number(
                text(styled, "A")
                    ?.getAttribute("transform")
                    ?.match(/translate\([-\d.]+[ ,]([-\d.]+)/)?.[1],
            ),
        ).toBeLessThan(
            Number(
                text(plain, "A")
                    ?.getAttribute("transform")
                    ?.match(/translate\([-\d.]+[ ,]([-\d.]+)/)?.[1],
            ),
        );
        // One dashed vertical gridline per category boundary, in the gridline color.
        const dashed = Array.from(styled?.querySelectorAll("path") ?? []).filter(path => path.getAttribute("stroke") === "#A5A5A5" && path.hasAttribute("stroke-dasharray"));

        expect(dashed.length).toBeGreaterThanOrEqual(3);
        expect(Array.from(plain?.querySelectorAll("path") ?? []).filter(path => path.hasAttribute("stroke-dasharray"))).toHaveLength(0);
    });

    test("normalizes the plot area fill, border and manual layout as PptxGenJS writes them", () => {
        const plain = normalizeChart("bar", categoricalData);

        // No plotArea: <a:noFill/> and <a:ln><a:noFill/></a:ln>; no layout: <c:layout/> (automatic).
        expect(plain.plotArea).toEqual({ fill: { visible: false, color: "transparent" as typeof plain.plotArea.borderColor }, borderVisible: false, borderColor: "#363636" as typeof plain.plotArea.borderColor, borderWidth: 1 });
        expect(plain.plotLayout).toBeUndefined();

        const styled = normalizeChart("line", categoricalData, { plotArea: { fill: { color: "DEEBF7", transparency: 50 }, border: { pt: 2, color: "7F7F7F" } }, layout: { x: 0.15, y: 0.1, w: 0.8, h: 0.6 } });

        expect(styled.plotArea.fill.visible).toBe(true);
        expect(styled.plotArea.fill.color).toBe(normalizeColor({ color: "DEEBF7", transparency: 50 }));
        expect(styled.plotArea).toMatchObject({ borderVisible: true, borderColor: "#7F7F7F", borderWidth: 2 });
        expect(styled.plotLayout).toEqual({ x: 0.15, y: 0.1, w: 0.8, h: 0.6 });
        // DEF_CHART_BORDER fills a border without pt or color; a fill without a color writes <a:noFill/>.
        expect(normalizeChart("bar", categoricalData, { plotArea: { border: {}, fill: { transparency: 50 } } }).plotArea).toMatchObject({ fill: { visible: false }, borderVisible: true, borderColor: "#363636", borderWidth: 1 });

        // Out-of-range values are dropped with PptxGenJS's warning, then `|| default`: w 0 is the full width.
        const warn = console.warn;
        const warnings: unknown[] = [];

        console.warn = (message: unknown) => warnings.push(message);

        try {
            expect(normalizeChart("bar", categoricalData, { layout: { x: 2, y: 0.1, w: 0, h: -1 } }).plotLayout).toEqual({ x: 0, y: 0.1, w: 1, h: 1 });
        } finally {
            console.warn = warn;
        }

        expect(warnings).toHaveLength(2);
        expect(String(warnings[0])).toContain("layout.x is 2");
    });

    test("rejects plot-area options the chart type does not draw and unknown plot-area keys", () => {
        expect(() => normalizeChart("pie", [categoricalData[0]!], { plotArea: { fill: { color: "FF0000" } } })).toThrow("plotArea[pie]");
        expect(() => normalizeChart("radar", categoricalData, { fill: "FF0000" } as PptxGenJS.IChartOpts)).toThrow("fill[radar]");
        expect(() => normalizeChart("bar", categoricalData, { plotArea: { border: { color: "FF0000", dashType: "dash" } as PptxGenJS.BorderProps } })).toThrow("plotArea.border.dashType");
        expect(() => normalizeChart("bar", categoricalData, { border: { color: "FF0000", dashType: "dash" } } as PptxGenJS.IChartOpts)).toThrow("border.dashType");
    });

    test("maps the deprecated top-level fill and border onto the plot area as PptxGenJS 4.0.1 does (#479)", () => {
        const chart = normalizeChart("bar", categoricalData, { fill: "F2F2F2", border: { pt: 2, color: "C00000" } } as PptxGenJS.IChartOpts);

        expect(chart.plotArea).toMatchObject({ fill: { visible: true, color: "#F2F2F2" }, borderVisible: true, borderColor: "#C00000", borderWidth: 2 });
        // A layout places pie, radar and bar3D charts too.
        for (const type of ["pie", "radar", "bar3D"] as const) {
            expect(normalizeChart(type, type === "pie" ? [categoricalData[0]!] : categoricalData, { layout: { x: 0.1, y: 0.1, w: 0.5, h: 0.5 } }).plotLayout).toEqual({ x: 0.1, y: 0.1, w: 0.5, h: 0.5 });
        }
    });

    test("draws the plot area fill and border on the plot rectangle and places it by the manual layout", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("bar", categoricalData, { x: 0, y: 0, w: 4, h: 3, plotArea: { fill: { color: "F2F2F2" }, border: { pt: 1, color: "7F7F7F" } }, layout: { x: 0.25, y: 0.1, w: 0.5, h: 0.6 } });
        const svg = renderedPage(presentation).querySelector(".slide-chart svg");
        const plot = Array.from(svg?.querySelectorAll("path") ?? []).find(path => path.getAttribute("fill") === "#F2F2F2");

        expect(plot?.getAttribute("stroke")).toBe("#7F7F7F");
        // The rectangle is the layout fractions of the 4in x 3in frame: from (0.25w, 0.1h), 0.5w wide and 0.6h high.
        const width = inchesToPixels(4);
        const height = inchesToPixels(3);

        expect(plot?.getAttribute("d")).toBe(`M${0.25 * width} ${0.1 * height}l${0.5 * width} 0l0 ${+(0.6 * height).toFixed(2)}l-${0.5 * width} 0Z`);
    });

    test("applies barGrouping per chart type as PptxGenJS writes <c:grouping> (#518)", () => {
        const grouping = (type: PptxGenJS.CHART_NAME, barGrouping?: string) => normalizeChart(type, categoricalData, barGrouping === undefined ? {} : { barGrouping }).grouping;

        // A line chart gets no <c:grouping>, so it is drawn standard and its automatic axis is not stacked.
        expect(grouping("line", "stacked")).toBe("standard");
        expect(grouping("line", "percentStacked")).toBe("standard");
        expect(normalizeChart("line", categoricalData, { barGrouping: "stacked" }).valueAxisMaximum).toBe(7);
        // An area chart is stacked only for "stacked"; percentStacked writes no <c:grouping>.
        expect(grouping("area", "stacked")).toBe("stacked");
        expect(grouping("area", "percentStacked")).toBe("standard");
        expect(grouping("area")).toBe("standard");
        // A bar chart turns standard (and anything else) into clustered; bar3D keeps standard and defaults to it.
        expect(grouping("bar")).toBe("clustered");
        expect(grouping("bar", "standard")).toBe("clustered");
        expect(grouping("bar", "percentStacked")).toBe("percentStacked");
        expect(grouping("bar3D")).toBe("standard");
        expect(grouping("bar3D", "clustered")).toBe("clustered");
        expect(grouping("radar", "stacked")).toBe("standard");
    });

    test("stacks each combo chart type as PptxGenJS writes its chart element's <c:grouping> (#529)", () => {
        const combo: PptxGenJS.IChartMulti[] = [
            { type: "bar", data: categoricalData, options: {} },
            { type: "line", data: [{ name: "Line", labels: ["A", "B", "C"], values: [6, 9, 7] }], options: {} },
            {
                type: "area",
                data: [
                    { name: "Area 1", labels: ["A", "B", "C"], values: [2, 3, 2] },
                    { name: "Area 2", labels: ["A", "B", "C"], values: [1, 2, 1] },
                ],
                options: {},
            },
        ];
        const stacks = (barGrouping?: string) => normalizeChart(combo, [], barGrouping === undefined ? {} : { barGrouping }, "powerpoint").series.map(series => series.stack);

        // The bars take the grouping, the areas stack only for "stacked", the line never does.
        expect(stacks("stacked")).toEqual(["stacked", "stacked", undefined, "stacked", "stacked"]);
        expect(stacks("percentStacked")).toEqual(["shareStacked", "shareStacked", undefined, undefined, undefined]);
        expect(normalizeChart("bar", categoricalData, { barGrouping: "percentStacked" }, "powerpoint").series[0]?.stack).toBe("percentStacked");
        expect(stacks()).toEqual([undefined, undefined, undefined, undefined, undefined]);
        expect(normalizeChart(combo, [], { barGrouping: "standard" }).grouping).toBe("standard");
        // The automatic axis covers the bar stack (6, 6, 9) and the line (6, 9, 7) apart: 9 is the largest, not 9 + 9.
        expect(normalizeChart(combo, [], { barGrouping: "stacked" }).valueAxisMaximum).toBe(10);
        // PowerPoint draws a combo's percent-stacked bars as shares of 1 on the axis the line and areas set (#537): 0..10.
        expect(normalizeChart(combo, [], { barGrouping: "percentStacked" }, "powerpoint").valueAxisMaximum).toBe(10);
        // LibreOffice 26.2 stacks them by value (comboPercentBarStack, #537).
        const libreOffice = normalizeChart(combo, [], { barGrouping: "percentStacked" }, "libreoffice");

        expect(libreOffice.series.map(series => series.stack)).toEqual(["stacked", "stacked", undefined, undefined, undefined]);
        expect(libreOffice.grouping).toBe("percentStacked");
        expect(libreOffice.valueAxisMaximum).toBe(10);
    });

    test("reads a combo chart type's own barGrouping, chartColors and line and bar options (#590)", () => {
        const line = { name: "Line", labels: ["A", "B", "C"], values: [6, 9, 7] };
        const combo = (barOptions: PptxGenJS.IChartOpts, lineOptions: PptxGenJS.IChartOpts = {}): PptxGenJS.IChartMulti[] => [
            { type: "bar", data: categoricalData, options: barOptions },
            { type: "line", data: [line], options: lineOptions },
        ];
        // PptxGenJS writes each type with `{ ...options, ...type.options }`: the bar type's barGrouping stacks its bars.
        const stacked = normalizeChart(combo({ barGrouping: "stacked" }), [], {});

        expect(stacked.series.map(series => series.stack)).toEqual(["stacked", "stacked", undefined]);
        expect(stacked.grouping).toBe("clustered");
        expect(stacked.valueAxisMaximum).toBe(10);
        expect(stacked.barOverlapPercent).toBeUndefined();
        // A type's own grouping replaces the chart's; a line's draws nothing.
        expect(normalizeChart(combo({ barGrouping: "clustered" }, { barGrouping: "stacked" }), [], { barGrouping: "stacked" }).series.map(series => series.stack)).toEqual([undefined, undefined, undefined]);
        // A stacked bar type keeps the label positions of a stack.
        expect(normalizeChart(combo({ barGrouping: "stacked", showValue: true, dataLabelPosition: "outEnd" }), [], {}).series[0]?.dataLabels?.position).toBe("outEnd");

        // Each type's chartColors color its own series, after the chart's palette, as in the PptxGenJS demo.
        const colored = normalizeChart(combo({ chartColors: ["CC0000", "33CC00"] }, { chartColors: ["0088CC"] }), [], { chartColors: ["111111"] });

        expect<string[]>(colored.colors).toEqual(["#111111", "#CC0000", "#33CC00", "#0088CC"]);
        expect<(string | undefined)[]>(colored.series.map(series => colored.colors[series.colorIndex])).toEqual(["#CC0000", "#33CC00", "#0088CC"]);
        // A single-series bar type with more than one color of its own fills each point from them.
        expect<string[] | undefined>(
            normalizeChart(
                [
                    { type: "bar", data: [categoricalData[0]!], options: { chartColors: ["CC0000", "33CC00"] } },
                    { type: "line", data: [line], options: {} },
                ],
                [],
                {},
            ).series[0]?.pointFills,
        ).toEqual(["#CC0000", "#33CC00", "#CC0000"]);

        // A line type's line options style its series; a bar type's gap and overlap are drawn chart-wide.
        const styled = normalizeChart(combo({ barGapWidthPct: 80, barOverlapPct: 20 }, { lineSize: 4, lineSmooth: true, lineDataSymbol: "square", lineDataSymbolSize: 9, lineDataSymbolLineColor: "FF0000" }), [], { lineSize: 1 });
        const lineStyle = styled.series[2]?.style;

        expect<unknown[]>([lineStyle?.lineSize, lineStyle?.lineSmooth, lineStyle?.lineSymbol, lineStyle?.lineSymbolSize, lineStyle?.lineSymbolLineColor]).toEqual([4, true, "square", 9, "#FF0000"]);
        expect([styled.lineSize, styled.series[0]?.style]).toEqual([1, undefined]);
        expect([styled.barGapWidthPercent, styled.barOverlapPercent]).toEqual([80, 20]);
        // A false secondaryCatAxis is the default.
        expect(() => normalizeChart(combo({}, { secondaryCatAxis: false }), [], {})).not.toThrow();
    });

    test("styles each combo chart type's series from its own options (#594)", () => {
        const line = { name: "Line", labels: ["A", "B", "C"], values: [6, 9, 7] };
        const otherLine = { name: "Other", labels: ["A", "B", "C"], values: [2, 4, 3] };
        // Two line types with their own line sizes and markers; LibreOffice 26.2.6 smooths both as the first says.
        const twoLines = normalizeChart(
            [
                { type: "bar", data: categoricalData, options: {} },
                { type: "line", data: [line], options: { lineSize: 1, lineSmooth: true } },
                { type: "line", data: [otherLine], options: { lineSize: 5, lineDataSymbol: "square" } },
            ],
            [],
            {},
        );

        expect(twoLines.series.map(series => [series.style?.lineSize, series.style?.lineSymbol, series.style?.lineSmooth])).toEqual([
            [undefined, undefined, undefined],
            [undefined, undefined, undefined],
            [1, "circle", true],
            [5, "square", true],
        ]);
        expect(twoLines.lineSmooth).toBe(true);
        // PowerPoint's PDF export smooths each line type as its own lineSmooth says (comboLineSmoothFromFirstLineType).
        const powerPointLines = normalizeChart(
            [
                { type: "bar", data: categoricalData, options: {} },
                { type: "line", data: [line], options: { lineSize: 1 } },
                { type: "line", data: [otherLine], options: { lineSize: 5, lineSmooth: true } },
            ],
            [],
            {},
            "powerpoint",
        );

        expect(powerPointLines.series.map(series => series.style?.lineSmooth)).toEqual([undefined, undefined, false, true]);

        // A bar type's own opacity (in percent; PptxGenJS writes no alpha for 0, so it is opaque) and outline; a type's
        // raw lineDataSymbolLineSize is EMU.
        const own = normalizeChart(
            [
                { type: "bar", data: categoricalData, options: { chartColorsOpacity: 40, dataBorder: { pt: 2, color: "FF0000" } } },
                { type: "line", data: [line], options: { lineDataSymbolLineSize: 12700 } },
            ],
            [],
            { chartColorsOpacity: 0 },
        );

        expect(own.colorOpacity).toBe(1);
        expect(own.series[0]?.style).toMatchObject({ colorOpacity: 0.4, dataBorder: { color: "#FF0000", width: 2 } });
        expect(own.series[2]?.style?.lineSymbolLineSize).toBe(1);

        // A single-series bar type's own invertedColors fill its negative points.
        const inverted = normalizeChart(
            [
                { type: "bar", data: [{ name: "Lorem", labels: ["A", "B", "C", "D"], values: [10, -8, 30, -15] }], options: { chartColors: ["4472C4", "ED7D31"], invertedColors: ["00B050", "7030A0"] } },
                { type: "line", data: [{ ...line, labels: ["A", "B", "C", "D"], values: [5, 15, 10, 20] }], options: {} },
            ],
            [],
            {},
        );

        expect<string[] | undefined>(inverted.series[0]?.pointFills).toEqual(["#4472C4", "#7030A0", "#4472C4", "#7030A0"]);

        // A type's dataBorder is written unchecked, so it needs its pt and color.
        try {
            normalizeChart(
                [
                    { type: "bar", data: categoricalData, options: { dataBorder: { color: "FF0000" } } as PptxGenJS.IChartOpts },
                    { type: "line", data: [line], options: {} },
                ],
                [],
                {},
            );
            throw new Error("Expected normalization to fail");
        } catch (error) {
            expect((error as UnsupportedChartError).unsupportedOptions).toEqual(["series[0].options.dataBorder[needs a positive pt and a color]"]);
        }
    });

    test("reads catAxes and a secondary category axis (#594, #603)", () => {
        const secondary: PptxGenJS.IChartMulti[] = [
            { type: "bar", data: categoricalData, options: {} },
            { type: "line", data: [{ name: "Line", labels: ["A", "B", "C"], values: [100, 300, 200] }], options: { secondaryValAxis: true, secondaryCatAxis: true } },
        ];
        // PptxGenJS writes the primary category axis from `{ ...options, ...catAxes[0] }`.
        const chart = normalizeChart(secondary, [], { valAxes: [{}, {}], catAxes: [{ catAxisLabelColor: "FF0000" }, {}] });

        expect<string>(chart.categoryAxis.labelColor).toBe("#FF0000");
        expect(chart.valueAxes).toHaveLength(2);
        expect(() => normalizeChart(secondary, [], { valAxes: [{}, {}], catAxes: [{}, {}, {}] })).toThrow("catAxes must have as many entries as valAxes");

        const rejections: [PptxGenJS.IChartOpts, string][] = [
            [{ valAxes: [{}, {}] }, "series[1].options.secondaryCatAxis[needs secondaryValAxis and catAxes[1]]"],
            [{ valAxes: [{}, {}], catAxes: [{}, { catAxisLineColor: "FF0000" }] }, "catAxes[1].catAxisLineColor"],
            [{ valAxes: [{}, {}], catAxes: [{ catAxisCrossesAt: 2 }, {}] }, "catAxes[0].catAxisCrossesAt"],
        ];

        for (const [options, rejected] of rejections) {
            try {
                normalizeChart(secondary, [], options);
                throw new Error("Expected normalization to fail");
            } catch (error) {
                expect((error as UnsupportedChartError).unsupportedOptions).toEqual([rejected]);
            }
        }

        // catAxes[1] keeps the chart-level catAxis* options under its own; PowerPoint draws it in a row of its own and
        // runs the line along it, LibreOffice draws it over the primary labels and keeps the line on the primary axis.
        const both = { valAxes: [{}, {}], catAxisLabelFontSize: 9, catAxes: [{}, { catAxisLabelColor: "C00000", catAxisOrientation: "maxMin", showCatAxisTitle: true, catAxisTitle: "Sec" }] } as PptxGenJS.IChartOpts;
        const powerpoint = normalizeChart(secondary, [], both, "powerpoint");
        const libreoffice = normalizeChart(secondary, [], both);

        expect(powerpoint.secondaryCategoryAxis).toMatchObject({ placement: "row", labelColor: "#C00000", labelFontSize: 9, reversed: true, showTitle: true, title: "Sec", labelPosition: "low" });
        expect(powerpoint.categoryAxis.reversed).toBeUndefined();
        expect(powerpoint.series.map(series => series.onSecondaryCategoryAxis)).toEqual([undefined, undefined, true]);
        expect(libreoffice.secondaryCategoryAxis?.placement).toBe("overlay");
        expect(libreoffice.series.map(series => series.onSecondaryCategoryAxis)).toEqual([undefined, undefined, undefined]);
        expect(normalizeChart(secondary, [], { valAxes: [{}, {}], catAxes: [{}, { catAxisLabelPos: "high" }] }, "powerpoint").secondaryCategoryAxis?.labelPosition).toBe("high");

        // The secondary row is measured under a column chart's labels, without a data table.
        expect(() => normalizeChart(secondary, [], { valAxes: [{}, {}], catAxes: [{}, {}], barDir: "bar" }, "powerpoint")).toThrow("secondaryCatAxis[barDir bar]");
        expect(() => normalizeChart(secondary, [], { valAxes: [{}, {}], catAxes: [{}, {}], showDataTable: true })).toThrow("secondaryCatAxis[showDataTable]");

        // The SVG: the secondary labels in their color, reversed, and its title; the primary labels below them on
        // PowerPoint, at their height on LibreOffice.
        const labels = (target: "libreoffice" | "powerpoint") => {
            const presentation = new PuppeteerGen(undefined, { target });

            const slide = presentation.addSlide();

            // PptxGenJS takes a combo chart's options second, where its declarations type the data.
            (slide.addChart as Function)(secondary, { x: 1, y: 1, w: 4, h: 3, ...both });
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");
            const texts = Array.from(svg?.querySelectorAll("text") ?? []);
            const y = (text: Element | undefined) => Number(/translate\([^ ]+ ([^)]+)\)/.exec(text?.getAttribute("transform") ?? "")?.[1]);
            const x = (text: Element) => Number(/translate\(([^ ]+) /.exec(text.getAttribute("transform") ?? "")?.[1]);
            // Left to right.
            const red = texts.filter(text => text.getAttribute("fill") === "#C00000").sort((a, b) => x(a) - x(b));
            const primary = texts.find(text => text.textContent === "A" && text.getAttribute("fill") !== "#C00000");

            return { red: red.map(text => text.textContent), title: texts.some(text => text.textContent === "Sec"), below: y(red[0]) - y(primary) };
        };
        const powerpointLabels = labels("powerpoint");
        const libreofficeLabels = labels("libreoffice");

        expect(powerpointLabels.red).toEqual(["C", "B", "A"]);
        expect(powerpointLabels.title).toBe(true);
        expect(powerpointLabels.below).toBeGreaterThan(5);
        expect(libreofficeLabels.red).toEqual(["C", "B", "A"]);
        expect(libreofficeLabels.title).toBe(true);
        expect(libreofficeLabels.below).toBeCloseTo(0, 0);
    });

    test("normalizes PptxGenJS bar3D defaults and bounded view options", () => {
        const chart = normalizeChart("bar3D", categoricalData, {
            bar3DShape: "cylinder",
            barGapDepthPct: 240,
            v3DPerspective: 90,
            v3DRAngAx: true,
            v3DRotX: -20,
            v3DRotY: 390,
        });

        expect(chart.grouping).toBe("standard");
        expect(chart.bar3DShape).toBe("cylinder");
        expect(chart.barGapDepthPercent).toBe(240);
        expect(chart.perspective3D).toBe(90);
        expect(chart.rightAngleAxes3D).toBe(true);
        expect(chart.rotationX3D).toBe(-20);
        expect(chart.rotationY3D).toBe(30);
        expect(normalizeChart("bar3D", categoricalData, { bar3DShape: "invalid" }).bar3DShape).toBe("box");
        // Stacked totals reach 9: a 3D axis takes no extra step near its data (#541).
        expect(normalizeChart("bar3D", categoricalData, { barGrouping: "stacked" }).valueAxisMaximum).toBe(9);
    });

    test("numbers a scatter chart's x axis as PowerPoint does, from 0 by the room its labels leave (scatterXAxisScale, #616)", () => {
        // The x axis's labels, the last ones ECharts writes (after the y axis's).
        const xLabels = (target: "libreoffice" | "powerpoint", x: number[], wide: boolean, count: number): string[] => {
            const presentation = new PuppeteerGen(undefined, { target });

            presentation.addSlide().addChart(
                "scatter",
                [
                    { name: "X", values: x },
                    { name: "Lorem", values: [2, 5, 3, 7].slice(0, x.length) },
                ],
                { ...(wide ? { x: 0.6, y: 0.9, w: 8.8, h: 4.5 } : { x: 0.4, y: 0.9, w: 4.4, h: 4.1 }), catAxisLabelFontFace: "Arial", valAxisLabelFontFace: "Arial", lineSize: 0 },
            );
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");

            return Array.from(svg?.querySelectorAll("text") ?? [], text => text.textContent ?? "").slice(-count);
        };

        // PowerPoint's PDF exports of these charts, 4.4 and 8.8 inches wide.
        expect(xLabels("powerpoint", [1, 2, 4, 5], false, 4)).toEqual(["0", "2", "4", "6"]);
        expect(xLabels("powerpoint", [1, 2, 4, 5], true, 7)).toEqual(["0", "1", "2", "3", "4", "5", "6"]);
        expect(xLabels("powerpoint", [10, 20, 30, 40], false, 6)).toEqual(["0", "10", "20", "30", "40", "50"]);
        expect(xLabels("powerpoint", [10, 20, 30, 40], true, 10)).toEqual(["0", "5", "10", "15", "20", "25", "30", "35", "40", "45"]);
        // Its labels in General, as both applications print them: 1000, not 1,000.
        expect(xLabels("powerpoint", [1000, 2000, 3000, 4000], false, 6)).toEqual(["0", "1000", "2000", "3000", "4000", "5000"]);
        expect(xLabels("powerpoint", [-5, -2, 1, 5], false, 5)).toEqual(["-10", "-5", "0", "5", "10"]);
        expect(xLabels("powerpoint", [0, 30, 60, 100], false, 4)).toEqual(["0", "50", "100", "150"]);
        // LibreOffice 26.2.6 keeps the scale over the x values (#548).
        expect(xLabels("libreoffice", [10, 20, 30, 40], false, 9)).toEqual(["5", "10", "15", "20", "25", "30", "35", "40", "45"]);
    });

    test("steps a bar3D value axis by the target's label heights per step (bar3DLabelHeightsPerStep, #541)", () => {
        const ticks = (target: "libreoffice" | "powerpoint", options: PptxGenJS.IChartOpts = {}, series: number[][] = [[3, 5, 4, 7]]): string[] => {
            const presentation = new PuppeteerGen(undefined, { target });
            const data = series.map((values, index) => ({ name: `S${index + 1}`, labels: ["Q1", "Q2", "Q3", "Q4"], values }));

            presentation.addSlide().addChart("bar3D", data, { x: 1, y: 1.3, w: 3.2, h: 3, showLegend: true, legendPos: "b", ...options });
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");

            return Array.from(svg?.querySelectorAll("text[text-anchor='end']") ?? [], text => text.textContent ?? "");
        };
        const arial = { w: 4.2, showLegend: false, valAxisLabelFontFace: "Arial", valAxisLabelFontSize: 10, catAxisLabelFontFace: "Arial", catAxisLabelFontSize: 10 };

        // The chartBar3DAxis fixture's first chart: LibreOffice 26.2.6 steps 0..7 by 1.
        expect(ticks("libreoffice")).toEqual(["0", "1", "2", "3", "4", "5", "6", "7"]);
        // PowerPoint's PDF export of synthetic 4.2in wide bar3D charts (#589): a padded scale with no padding, 3, 5, 4, 7
        // and 2, 4, 6, 3 0..7 by 1 on a 4.5in chart, 3, 9, 16, 12 and 5, 7, 10, 4 0..20 by 5 on a 3in one, 950..1000
        // 920..1000 by 10 (the minimum dropped half the range) and 3, -2, 4, 7 -2..7 by 1 on 4.5in.
        expect(
            ticks("powerpoint", { ...arial, h: 4.5, barGrouping: "standard" }, [
                [3, 5, 4, 7],
                [2, 4, 6, 3],
            ]),
        ).toEqual(["0", "1", "2", "3", "4", "5", "6", "7"]);
        expect(
            ticks("powerpoint", { ...arial, h: 3, barGrouping: "standard" }, [
                [3, 9, 16, 12],
                [5, 7, 10, 4],
            ]),
        ).toEqual(["0", "5", "10", "15", "20"]);
        expect(ticks("powerpoint", { ...arial, h: 3 }, [[950, 970, 1000, 985]])).toEqual(["920", "930", "940", "950", "960", "970", "980", "990", "1000"]);
        expect(ticks("powerpoint", { ...arial, h: 4.5 }, [[3, -2, 4, 7]])).toEqual(["-2", "-1", "0", "1", "2", "3", "4", "5", "6", "7"]);
        // shortcut: this projection is longer than PowerPoint's on short charts (#541), so the fixture's chart, which
        // PowerPoint's export steps 0..10 by 5 on an 83px axis, steps 0..8 by 2 here; fix with the projection.
        expect(ticks("powerpoint")).toEqual(["0", "2", "4", "6", "8"]);
        // valAxisMajorUnit sets the step whatever the target.
        expect(ticks("powerpoint", { valAxisMajorUnit: 1 })).toEqual(["0", "1", "2", "3", "4", "5", "6", "7"]);
    });

    test("steps a 2D value axis by the room its labels have on the target (valueAxisStepRoom, #585, #589)", () => {
        const scale = (target: "libreoffice" | "powerpoint", type: "bar" | "line", values: number[], options: PptxGenJS.IChartOpts): string => {
            const presentation = new PuppeteerGen(undefined, { target });

            presentation.addSlide().addChart(type, [{ name: "Lorem", labels: ["A", "B", "C", "D"], values }], { x: 0.3, y: 0.7, w: 6, valAxisLabelFontFace: "Arial", valAxisLabelFontSize: 10, ...options });
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");
            const anchor = options.barDir === "bar" ? "middle" : "end";
            const labels = Array.from(svg?.querySelectorAll(`text[text-anchor='${anchor}']`) ?? [], text => text.textContent ?? "").filter(text => /^-?[\d.]+$/.test(text));

            return `${labels[0]}..${labels.at(-1)} (${labels.length})`;
        };

        expect(normalizeChart("bar", [{ name: "Lorem", labels: ["A"], values: [3] }]).valueAxes[0]?.automaticScale).toEqual({ low: 3, high: 3, fixedMinimum: false, fixedMaximum: false });
        expect(normalizeChart("bar", [{ name: "Lorem", labels: ["A"], values: [3] }], { valAxisMajorUnit: 1 }).valueAxes[0]?.automaticScale).toBeUndefined();
        // LibreOffice 26.2: 3, 5, 4, 7 at 10pt runs 0..8 by 2 on a 73pt axis and by 1 on an 86.5pt one.
        expect(scale("libreoffice", "bar", [3, 5, 4, 7], { h: 1.4 })).toBe("0..8 (5)");
        expect(scale("libreoffice", "bar", [3, 5, 4, 7], { h: 1.7 })).toBe("0..8 (9)");
        // PowerPoint's PDF export: 3, -2, 4, 7 runs -3..8 by 1 on a 181pt axis, -5..10 by 5 on a 66pt one, as a line too.
        expect(scale("powerpoint", "bar", [3, -2, 4, 7], { h: 3 })).toBe("-3..8 (12)");
        expect(scale("powerpoint", "bar", [3, -2, 4, 7], { h: 1.4 })).toBe("-5..10 (4)");
        expect(scale("powerpoint", "line", [3, -2, 4, 7], { h: 4.4 })).toBe("-3..8 (12)");
        // A horizontal axis: 3, 9, 16, 12 on a 4in chart runs 0..18 by 2 in LibreOffice, 0..20 by 5 in PowerPoint.
        expect(scale("libreoffice", "bar", [3, 9, 16, 12], { h: 4.4, w: 4, barDir: "bar" })).toBe("0..18 (10)");
        expect(scale("powerpoint", "bar", [3, 9, 16, 12], { h: 4.4, w: 4, barDir: "bar" })).toBe("0..20 (5)");
    });

    test("steps a bar3D value axis as LibreOffice does, by valAxisMajorUnit when set (#541)", () => {
        const one: PptxGenJS.OptsChartData[] = [{ name: "Lorem", labels: ["A", "B", "C", "D"], values: [3, 5, 4, 7] }];
        const ticks = (options: PptxGenJS.IChartOpts, h = 4.8): string[] => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("bar3D", one, { x: 0.5, y: 0.6, w: 9, h, fontFace: "Arial", valAxisLabelFontSize: 14, ...options });
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");

            return Array.from(svg?.querySelectorAll("text[text-anchor='end']") ?? [], text => text.textContent ?? "");
        };

        expect(normalizeChart("bar3D", one).valueAxes[0]?.automaticScale).toEqual({ low: 0, high: 7, fixedMinimum: false, fixedMaximum: false });
        // LibreOffice 26.2.6 on the same charts: 0..7 by 1, 0..8 by 2, 0..7.5 by 2.5, 0..10 by 2 and 2..9 by 1.
        expect(ticks({})).toEqual(["0", "1", "2", "3", "4", "5", "6", "7"]);
        expect(ticks({ valAxisMajorUnit: 2 })).toEqual(["0", "2", "4", "6", "8"]);
        expect(ticks({ valAxisMajorUnit: 2.5 })).toEqual(["0", "2.5", "5", "7.5"]);
        expect(ticks({ valAxisMaxVal: 10 })).toEqual(["0", "2", "4", "6", "8", "10"]);
        expect(ticks({ valAxisMinVal: 2, valAxisMaxVal: 9 })).toEqual(["2", "3", "4", "5", "6", "7", "8", "9"]);
        // The labels take the value axis's font and format.
        const presentation = new PuppeteerGen();

        presentation.addSlide().addChart("bar3D", one, { x: 0, y: 0, w: 6, h: 4, valAxisLabelFontSize: 9, valAxisLabelColor: "FF0000", valAxisLabelFormatCode: "0.0" });
        const label = renderedPage(presentation).querySelector(".slide-chart svg text[text-anchor='end']");

        expect([label?.getAttribute("font-size"), label?.getAttribute("fill"), label?.textContent]).toEqual(["12", "#FF0000", "0.0"]);
    });

    test("names the bar3D depth slots only when there are several, or a single series where the target does (#541)", () => {
        const depthLabels = (data: PptxGenJS.OptsChartData[], options: PptxGenJS.IChartOpts = {}, target: "libreoffice" | "powerpoint" = "libreoffice"): (string | null)[] => {
            const presentation = new PuppeteerGen(undefined, { target });

            presentation.addSlide().addChart("bar3D", data, { x: 0, y: 0, w: 6, h: 4, ...options });

            return Array.from(renderedPage(presentation).querySelectorAll(".bar3d-series-label"), label => label.textContent);
        };

        expect(depthLabels(categoricalData)).toEqual(["First", "Second"]);
        expect(depthLabels([categoricalData[0]!])).toEqual([]);
        expect(depthLabels(categoricalData, { barGrouping: "stacked" })).toEqual([]);
        // PowerPoint's PDF export of a synthetic deck names a single series there (bar3DNamesSingleSeries); a stacked
        // chart keeps LibreOffice's none, unmeasured on PowerPoint.
        expect(depthLabels([categoricalData[0]!], {}, "powerpoint")).toEqual(["First"]);
        expect(depthLabels(categoricalData, { barGrouping: "stacked" }, "powerpoint")).toEqual([]);
    });

    test("renders all six bar3D shapes as deterministic inline SVG", () => {
        const shapes = ["box", "cylinder", "cone", "coneToMax", "pyramid", "pyramidToMax"] as const;

        for (const shape of shapes) {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("bar3D", categoricalData, {
                x: 0,
                y: 0,
                w: 5,
                h: 3,
                bar3DShape: shape,
                showLegend: true,
            });
            const svg = renderedPage(presentation).querySelector(`svg[data-bar3d-shape="${shape}"]`);

            expect(svg).not.toBeNull();
            expect(svg?.querySelectorAll(".bar3d-mark")).toHaveLength(6);
            expect(svg?.textContent).toContain("First");
        }
    });

    test("escapes the bar3D title, series names and model strings it writes into attributes (#283)", () => {
        const tricky = `a<b & "c" 'd'`;
        const presentation = new PuppeteerGen();

        presentation.addSlide().addChart(
            "bar3D",
            [
                { ...categoricalData[0], name: tricky },
                { ...categoricalData[1], name: tricky },
            ],
            { x: 0, y: 0, w: 5, h: 3, showTitle: true, title: tricky, showLegend: true },
        );
        const svg = renderedPage(presentation).querySelector(".slide-chart svg");

        expect(svg?.querySelector(".bar3d-title")?.textContent).toBe(tricky);
        expect(Array.from(svg?.querySelectorAll(".bar3d-series-label, .bar3d-legend") ?? [], label => label.textContent)).toEqual([tricky, tricky, tricky, tricky]);
        expect(svg?.querySelectorAll("b")).toHaveLength(0);

        // A model from another producer is not normalized, so its colors and shape name reach the SVG as given.
        const injected = `red"><script>alert(1)</script><x y="`;
        const model = presentation.getModel();
        const chart = model.slides[0]?.objects[0];

        if (chart?.kind !== "chart") {
            throw new Error("expected the chart object");
        }

        Object.assign(chart.chart, { titleColor: injected, legendColor: injected, gridLineColor: injected, colors: [injected], bar3DShape: injected });
        const page = new JSDOM(unresolvedModelHtml(model)).window.document;

        expect(page.querySelectorAll("script, x")).toHaveLength(0);
        expect(page.querySelector(".bar3d-title")?.getAttribute("fill")).toBe(injected);
        expect(page.querySelector(".bar3d-legend")?.getAttribute("fill")).toBe(injected);
        expect(page.querySelector("svg[data-chart-type='bar3D']")?.getAttribute("data-bar3d-shape")).toBe(injected);
    });

    test("projects bar3D through the view3D camera PptxGenJS writes", () => {
        const walls = (options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("bar3D", categoricalData, { x: 0, y: 0, w: 6, h: 3.3, ...options });
            const page = renderedPage(presentation);
            const corners = (element: Element | undefined) => (element?.getAttribute("points") ?? "").split(" ").map(pair => pair.split(",").map(Number) as [number, number]);
            const side = page.querySelectorAll(".slide-chart .bar3d-wall")[1];
            const xs = Array.from(page.querySelectorAll(".slide-chart .bar3d-wall, .slide-chart .bar3d-floor")).flatMap(polygon => corners(polygon).map(([x]) => x));
            // Side wall: front bottom, back bottom, back top, front top.
            const [frontBottom, backBottom, backTop, frontTop] = corners(side);

            return { sideX: Math.min(...corners(side).map(([x]) => x)), frontEdge: (frontBottom?.[1] ?? 0) - (frontTop?.[1] ?? 0), backEdge: (backBottom?.[1] ?? 0) - (backTop?.[1] ?? 0), left: Math.min(...xs), right: Math.max(...xs) };
        };
        // `<c:view3D><c:rotX val="30"/><c:rotY val="30"/><c:rAngAx val="0"/><c:perspective val="30"/></c:view3D>`: a 15° field
        // of view makes the far edge of the side wall shorter; the box is centered with 12.5% of the width on each side.
        const perspective = walls({});

        expect(perspective.backEdge).toBeLessThan(perspective.frontEdge * 0.97);
        expect(perspective.left).toBeCloseTo(576 * 0.125, 0);
        expect(perspective.right).toBeCloseTo(576 * 0.875, 0);
        // Right-angle axes ignore the perspective: parallel edges keep their length.
        const parallel = walls({ v3DRAngAx: true });

        expect(parallel.backEdge).toBeCloseTo(parallel.frontEdge, 1);
        // Turning the scene the other way shows the other side wall.
        const turned = walls({ v3DRotY: 330 });

        expect(turned.sideX).toBeGreaterThan(576 / 2);
        expect(perspective.sideX).toBeLessThan(576 / 2);
    });

    test("sizes the automatic plot-area bands from the text, not the chart height, as LibreOffice lays them out (#528)", () => {
        const chartPlot = (hInches: number, options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("bar", categoricalData, { x: 0, y: 0, w: 5, h: hInches, ...options });
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");
            const rows: number[] = [];
            const columns: number[] = [];

            // The gridlines and axis lines: horizontal ones give a column chart's plot top and bottom, vertical ones a
            // horizontal bar chart's top.
            for (const path of svg?.querySelectorAll("path") ?? []) {
                const segment = /^M([\d.]+) ([\d.]+)L([\d.]+) ([\d.]+)$/.exec(path.getAttribute("d") ?? "");

                if (!segment) {
                    continue;
                }

                const [, x1, y1, x2, y2] = segment.map(Number);

                if (y1 === y2 && Number(x2) - Number(x1) > 100) {
                    rows.push(Number(y1));
                }

                if (x1 === x2 && Number(y2) - Number(y1) > 100) {
                    columns.push(Number(y1));
                }
            }

            // The first legend key's center.
            const key = svg?.querySelector('path[ecmeta_ssr_type="legend"][d^="M0 0"]');
            const keyY = Number(/translate\([-\d.]+ ([-\d.]+)\)/.exec(key?.getAttribute("transform") ?? "")?.[1]);
            const keySize = Math.round(pointsToPixels(10) * PPTX_DEFAULTS.chart.legendKeyFontRatio);

            return { top: Math.min(...rows), bottom: hInches * 96 - Math.max(...rows), columnTop: Math.min(...columns), keyCenter: keyY + keySize / 2 };
        };
        const band = PPTX_DEFAULTS.chart.plotBand;
        const legendPx = pointsToPixels(10);
        // LibreOffice 26.2's bands (#528): the bottom axis's 1.5mm ticks, 1mm label spacing and label line; a top or
        // bottom legend's row of 1.58 legend font sizes and 6pt; a title's gap, 1.7 title font sizes and 4.25pt; half the
        // 12pt value label above a column chart. Within 1px: zrender snaps lines to half pixels.
        const labelBand = pointsToPixels(band.tickLengthPt) + pointsToPixels(band.labelSpacingPt) + 16 * band.labelLineEm;
        const legendBand = legendPx * band.legendBandEm + pointsToPixels(band.legendBandPt);
        const rowCenter = (legendPx * band.legendBandEm) / 2;

        for (const h of [2.6, 4.1]) {
            const gap = h * 96 * PPTX_DEFAULTS.chart.plotAreaGapRatio;
            const titleBand = gap + 24 * PPTX_DEFAULTS.chart.pieLayout.titleBandEm + pointsToPixels(band.titleBandPt);
            const plain = chartPlot(h, {});
            const bottomLegend = chartPlot(h, { showLegend: true, legendPos: "b" });
            const topLegend = chartPlot(h, { showLegend: true, legendPos: "t" });
            const titled = chartPlot(h, { showTitle: true, title: "Lorem", showLegend: true, legendPos: "t" });

            expect(Math.abs(plain.top - (gap + 8))).toBeLessThanOrEqual(1);
            expect(Math.abs(plain.bottom - (gap + labelBand))).toBeLessThanOrEqual(1);
            expect(Math.abs(bottomLegend.bottom - (gap + legendBand + labelBand))).toBeLessThanOrEqual(1);
            expect(Math.abs(topLegend.top - (gap + legendBand + 8))).toBeLessThanOrEqual(1);
            expect(Math.abs(titled.top - (gap + titleBand + legendBand + 8))).toBeLessThanOrEqual(1);
            // A horizontal bar chart's top is a category band: no value label above it.
            expect(Math.abs(chartPlot(h, { barDir: "bar" }).columnTop - gap)).toBeLessThanOrEqual(1);
            // The legend entries are centered in their row: 6pt below the title band, or legendBottomPt above the bottom edge.
            expect(topLegend.keyCenter).toBeCloseTo(pointsToPixels(band.legendBandPt) + rowCenter, 1);
            expect(titled.keyCenter).toBeCloseTo(titleBand + pointsToPixels(band.legendBandPt) + rowCenter, 1);
            expect(bottomLegend.keyCenter).toBeCloseTo(h * 96 - pointsToPixels(band.legendBottomPt) - rowCenter, 1);
        }

        // Without outside ticks on the bottom axis there is no tick band.
        const noTicks = chartPlot(4.1, { barDir: "bar", valAxisMajorTickMark: "none" });

        expect(Math.abs(noTicks.bottom - (4.1 * 96 * PPTX_DEFAULTS.chart.plotAreaGapRatio + labelBand - pointsToPixels(band.tickLengthPt)))).toBeLessThanOrEqual(1);
    });

    test("places a top-right legend as each target draws it: beside the plot in LibreOffice, above it in PowerPoint (#536)", () => {
        const data = ["Lorem", "Ipsum", "Dolor"].map((name, index) => ({ name, labels: ["Q1", "Q2"], values: [index + 1, index + 2] }));
        const layout = (presentation: PuppeteerGen, options: PptxGenJS.IChartOpts) => {
            presentation.addSlide().addChart("bar", structuredClone(data), { x: 0, y: 0, w: 4.4, h: 4.1, ...options });
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");
            const rows: number[] = [];
            const ends: number[] = [];

            // The gridlines: the plot area's top and right edge.
            for (const path of svg?.querySelectorAll("path") ?? []) {
                const segment = /^M([\d.]+) ([\d.]+)L([\d.]+) ([\d.]+)$/.exec(path.getAttribute("d") ?? "");

                if (segment && segment[2] === segment[4] && Number(segment[3]) - Number(segment[1]) > 100) {
                    rows.push(Number(segment[2]));
                    ends.push(Number(segment[3]));
                }
            }

            const keys: { x: number; y: number }[] = [];

            for (const key of svg?.querySelectorAll('path[ecmeta_ssr_type="legend"][d^="M0 0"]') ?? []) {
                const translate = /translate\(([-\d.]+) ([-\d.]+)\)/.exec(key.getAttribute("transform") ?? "");

                keys.push({ x: Number(translate?.[1]), y: Number(translate?.[2]) });
            }

            return { top: Math.min(...rows), right: Math.max(...ends), keys };
        };
        const legendPx = pointsToPixels(10);
        const keySize = Math.round(legendPx * PPTX_DEFAULTS.chart.legendKeyFontRatio);
        // LibreOffice 26.2: in the chart's top corner, the first key centered 0.83 legend font sizes down, one entry every
        // 1.43; the plot keeps its top and its side band.
        const libreOffice = layout(new PuppeteerGen(), { showLegend: true, legendPos: "tr" });
        const libreOfficeRight = layout(new PuppeteerGen(), { showLegend: true, legendPos: "r" });
        const plain = layout(new PuppeteerGen(), {});
        const vertical = PPTX_DEFAULTS.chart.verticalLegend;

        expect(libreOffice.keys[0]?.y ?? NaN).toBeCloseTo(vertical.firstKeyCenterEm * legendPx - keySize / 2, 1);
        expect((libreOffice.keys[2]?.y ?? NaN) - (libreOffice.keys[0]?.y ?? NaN)).toBeCloseTo(2 * vertical.pitchEm * legendPx, 1);
        expect((libreOfficeRight.keys[1]?.y ?? NaN) - (libreOfficeRight.keys[0]?.y ?? NaN)).toBeCloseTo(vertical.pitchEm * legendPx, 1);
        expect(libreOffice.top).toBe(plain.top);
        expect(libreOffice.right).toBe(libreOfficeRight.right);
        expect(libreOffice.right).toBeLessThan(plain.right);

        // PowerPoint (its PDF export of a 4.4 x 4.1in stacked column chart): the first key's top 16.6px down, the last
        // key's bottom 70.6px, the top gridline 87.6px (14.6px without the legend), the plot area its full width.
        const powerPoint = layout(new PuppeteerGenPowerPoint(), { showLegend: true, legendPos: "tr", barGrouping: "stacked" });
        const powerPointPlain = layout(new PuppeteerGenPowerPoint(), { barGrouping: "stacked" });

        expect(Math.abs((powerPoint.keys[0]?.y ?? NaN) - 16.6)).toBeLessThan(0.6);
        expect(Math.abs((powerPoint.keys[2]?.y ?? NaN) + keySize - 70.6)).toBeLessThan(0.6);
        expect(Math.abs(powerPoint.top - 87.6)).toBeLessThanOrEqual(1.5);
        // The top gridline without a legend is a 1px line snapped to 16.5px, a pixel below PowerPoint's (#528).
        expect(Math.abs(powerPointPlain.top - 14.6)).toBeLessThanOrEqual(2);
        expect(powerPoint.right).toBe(powerPointPlain.right);
        // The model carries the target's placement, so another renderer draws it the same.
        expect(normalizeChart("bar", data, {}).topRightLegend).toEqual({ kind: "beside" });
    });

    test("reserves a side legend's band beside the plot area, as LibreOffice narrows it (#517)", () => {
        const width = inchesToPixels(4.4);
        // The gridlines' horizontal extent: the plot area's left and right edges.
        const plot = (type: "bar" | "line", options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart(type, categoricalData, { x: 0, y: 0, w: 4.4, h: 4.1, showLegend: true, ...options });
            const xs: number[] = [];

            for (const path of renderedPage(presentation).querySelectorAll(".slide-chart path")) {
                const match = /^M([\d.]+) ([\d.]+)L([\d.]+) \2$/.exec(path.getAttribute("d") ?? "");

                if (match) {
                    xs.push(Number(match[1]), Number(match[3]));
                }
            }

            return { left: Math.min(...xs), right: Math.max(...xs) };
        };
        // Two insets (9.75pt), the key, its 5px gap and the widest entry, measured as ECharts lays the legend out.
        const legendPx = pointsToPixels(PPTX_DEFAULTS.chart.legendFontSizePt);
        const inset = pointsToPixels(PPTX_DEFAULTS.chart.pieLayout.insetPt);
        const widest = format.getTextRect("Second", `${legendPx}px ${PPTX_DEFAULTS.theme.bodyFontFace}`).width;
        const band = 2 * inset + Math.round(legendPx * PPTX_DEFAULTS.chart.legendKeyFontRatio) + 5 + widest;
        const bottom = plot("bar", { legendPos: "b" });

        expect(bottom.right).toBeCloseTo(width * 0.98, 0);
        expect(plot("bar", { legendPos: "r" }).right).toBeCloseTo(width - band, 0);
        expect(plot("bar", { legendPos: "tr" }).right).toBeCloseTo(width - band, 0);
        // The default legend position is the right.
        expect(plot("bar", {}).right).toBeCloseTo(width - band, 0);
        // A line key keeps ECharts' 25px item width.
        expect(plot("line", { legendPos: "r" }).right).toBeCloseTo(width - (band - Math.round(legendPx * PPTX_DEFAULTS.chart.legendKeyFontRatio) + 25), 0);
        // A left legend moves the value axis gutter by the band less one inset; the right edge keeps its 2%.
        const left = plot("bar", { legendPos: "l" });

        expect(left.left).toBeCloseTo(bottom.left + band - inset, 0);
        expect(left.right).toBeCloseTo(width * 0.98, 0);
    });

    test("draws a point marker as PptxGenJS's c:marker writes it: size in points, its outline, none on area series", () => {
        const markers = (type: "area" | "line" | "radar" | "scatter", options: PptxGenJS.IChartOpts = {}, target: "libreoffice" | "powerpoint" = "libreoffice") => {
            const presentation = new PuppeteerGen(undefined, { target });

            presentation.addSlide().addChart(type, categoricalData, { x: 0, y: 0, w: 5, h: 3, ...options });

            return Array.from(renderedPage(presentation).querySelectorAll('.slide-chart path[ecmeta_ssr_type="chart"][ecmeta_data_index]'));
        };
        const lineMarkers = markers("line", {}, "powerpoint");
        // The default 6pt circle: a unit circle scaled to a 4px radius (8px across).
        const scale = pointsToPixels(PPTX_DEFAULTS.chart.lineDataSymbolSize) / 2;

        expect(lineMarkers).toHaveLength(6);
        expect(lineMarkers[0]?.getAttribute("transform")).toStartWith(`matrix(${scale},0,0,${scale},`);
        // PptxGenJS: `<a:ln w="${lineDataSymbolLineSize || 0.75}"><a:solidFill>${lineDataSymbolLineColor || seriesColor}`.
        expect(lineMarkers[0]?.getAttribute("stroke")).toBe(lineMarkers[0]?.getAttribute("fill") ?? "fill");
        // ECharts writes the stroke inside the scaled transform.
        expect(Number(lineMarkers[0]?.getAttribute("stroke-width")) * scale).toBe(pointsToPixels(PPTX_DEFAULTS.chart.lineDataSymbolLineSizePt));
        expect(normalizeChart("line", categoricalData, { lineDataSymbolLineSize: 0 }).lineSymbolLineSize).toBe(PPTX_DEFAULTS.chart.lineDataSymbolLineSizePt);

        const outlined = markers("line", { lineDataSymbolLineColor: "000000", lineDataSymbolLineSize: 1.5 }, "powerpoint");

        expect(outlined[0]?.getAttribute("stroke")).toBe("#000000");
        expect(Number(outlined[0]?.getAttribute("stroke-width")) * scale).toBe(pointsToPixels(1.5));
        expect(markers("scatter", { lineDataSymbolLineColor: "000000" }, "powerpoint")[0]?.getAttribute("stroke")).toBe("#000000");
        expect(markers("radar", { radarStyle: "marker", lineDataSymbolLineColor: "000000" }, "powerpoint")[0]?.getAttribute("stroke")).toBe("#000000");
        // LibreOffice draws no marker outline (drawsChartMarkerOutline).
        expect(markers("line", { lineDataSymbolLineColor: "000000", lineDataSymbolLineSize: 1.5 })[0]?.getAttribute("stroke-width")).toBe("0");
        // LibreOffice's shapes: `dot` a right-pointing triangle, `dash` a bar 0.2 of the marker size tall.
        expect(markers("line", { lineDataSymbol: "dot" })[0]?.getAttribute("d")).toBe("M-1 -1L1 0L-1 1Z");
        expect(markers("line", { lineDataSymbol: "dash" })[0]?.getAttribute("transform")).toStartWith(`matrix(${scale},0,0,${scale * 0.2},`);
        expect(markers("line", { lineDataSymbol: "square" })).toHaveLength(6);
        expect(markers("line", { lineDataSymbol: "none" })).toHaveLength(0);
        expect(markers("area")).toHaveLength(0);
        expect(markers("radar", { radarStyle: "marker" })).toHaveLength(6);
        expect(markers("radar")).toHaveLength(0);
        expect(markers("scatter", { lineDataSymbol: "none" })).toHaveLength(0);
    });

    test("fills an area series at chartColorsOpacity with no edge line, as PptxGenJS writes its c:spPr (#520)", () => {
        const areaPaths = (options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("area", categoricalData, { x: 0, y: 0, w: 5, h: 3, chartColors: ["4472C4", "ED7D31"], lineSize: 3, ...options });

            return Array.from(renderedPage(presentation).querySelectorAll(".slide-chart path")).filter(path => path.getAttribute("fill") === "#4472C4" || path.getAttribute("stroke") === "#4472C4");
        };
        // The fill at 100% by default; the series line ECharts draws along the top has no width (no <a:ln>, even with lineSize).
        const [fill, edge] = areaPaths({});

        expect(fill?.getAttribute("fill")).toBe("#4472C4");
        expect(fill?.getAttribute("fill-opacity")).toBeNull();
        expect(edge?.getAttribute("stroke-width")).toBe("0");
        // `<a:alpha val="${chartColorsOpacity * 1000}"/>`.
        expect(areaPaths({ chartColorsOpacity: 40 })[0]?.getAttribute("fill-opacity")).toBe("0.4");
    });

    test("draws a standard area chart's first series in front where the target does (firstAreaSeriesInFront, #533)", () => {
        const drawn = (options: PptxGenJS.IChartOpts, target: "libreoffice" | "powerpoint" = "libreoffice") => {
            const presentation = new PuppeteerGen(undefined, { target });

            presentation.addSlide().addChart("area", categoricalData, { x: 0, y: 0, w: 5, h: 3, chartColors: ["4472C4", "ED7D31"], ...options });
            const object = presentation.getModel().slides[0]?.objects[0];
            // The series fills in paint order: the last is in front.
            const fills: string[] = [];

            for (const path of renderedPage(presentation).querySelectorAll(".slide-chart path")) {
                const fill = path.getAttribute("fill");

                if ((fill === "#4472C4" || fill === "#ED7D31") && path.getAttribute("d")?.endsWith("Z")) {
                    fills.push(fill);
                }
            }

            return { fills, firstSeriesInFront: object?.kind === "chart" ? object.chart.firstSeriesInFront : undefined };
        };

        // LibreOffice: the overlapping series are drawn from the last, so the first is in front.
        expect(drawn({})).toEqual({ fills: ["#ED7D31", "#4472C4"], firstSeriesInFront: true });
        // PowerPoint draws them in series order, so the last is in front.
        expect(drawn({}, "powerpoint")).toEqual({ fills: ["#4472C4", "#ED7D31"], firstSeriesInFront: undefined });
        // A stacked chart's series do not overlap and keep the series order, which is also its stacking order.
        expect(drawn({ barGrouping: "stacked" })).toEqual({ fills: ["#4472C4", "#ED7D31"], firstSeriesInFront: undefined });
    });

    test("joins a scatter series' points with its line, as PptxGenJS's lineMarker style writes it", () => {
        const xyData: PptxGenJS.OptsChartData[] = [
            { name: "X", values: [1, 2, 3, 4] },
            { name: "Lorem", values: [2, 4, 3, 5] },
            { name: "Ipsum", values: [1, 2, 4, 2] },
        ];
        // The stroked, unfilled series paths (axis and grid lines are gray).
        const seriesLines = (options: PptxGenJS.IChartOpts = {}) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("scatter", xyData, { x: 0, y: 0, w: 5, h: 3, chartColors: ["4472C4", "ED7D31"], ...options });
            const lines = [];

            for (const path of renderedPage(presentation).querySelectorAll('.slide-chart path[fill="none"]')) {
                if (path.getAttribute("stroke") === "#4472C4" || path.getAttribute("stroke") === "#ED7D31") {
                    lines.push(path);
                }
            }

            return lines;
        };
        // PptxGenJS: `<a:ln w="${valToPts(lineSize)}">` (2pt by default), `<a:ln><a:noFill/></a:ln>` at 0, `<c:smooth>`.
        const lines = seriesLines();

        expect(lines).toHaveLength(2);
        expect(lines[0]?.getAttribute("d")).toMatch(/^M[\d.]+ [\d.]+(L[\d.]+ [\d.]+){3}$/);
        expect(Number(lines[0]?.getAttribute("stroke-width"))).toBe(pointsToPixels(PPTX_DEFAULTS.chart.lineSizePt));
        expect(Number(seriesLines({ lineSize: 4 })[0]?.getAttribute("stroke-width"))).toBe(pointsToPixels(4));
        expect(seriesLines({ lineSmooth: true })[0]?.getAttribute("d")).toContain("C");
        // Without markers the lines remain.
        expect(seriesLines({ lineDataSymbol: "none" })).toHaveLength(2);

        // The markers keep their outline in the series color.
        const presentation = new PuppeteerGen();

        presentation.addSlide().addChart("scatter", xyData, { x: 0, y: 0, w: 5, h: 3, chartColors: ["4472C4", "ED7D31"] });
        const marker = renderedPage(presentation).querySelector('.slide-chart path[ecmeta_ssr_type="chart"][ecmeta_series_index="1"][ecmeta_data_index]');

        expect([marker?.getAttribute("fill"), marker?.getAttribute("stroke")]).toEqual(["#ED7D31", "#ED7D31"]);

        for (const line of seriesLines({ lineSize: 0 })) {
            expect(line.getAttribute("stroke-width")).toBe("0");
        }
    });

    test("scales a scatter chart's axes as LibreOffice does, without ticks, the y axis by the value-axis options (#519)", () => {
        const xyData: PptxGenJS.OptsChartData[] = [
            { name: "X", values: [0, 2, 4, 6, 7.6] },
            { name: "Lorem", values: [3, 5, 4, 6.4, 5.5] },
        ];
        // The gridlines: vertical ones at each x step, horizontal ones at each y step.
        const steps = (options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("scatter", structuredClone(xyData), { x: 0, y: 0, w: 5, h: 3, ...options });
            const columns: number[] = [];
            const rows: number[] = [];
            const labels: string[] = [];
            let ticks = 0;

            for (const path of renderedPage(presentation).querySelectorAll(".slide-chart path")) {
                const segment = /^M([\d.]+) ([\d.]+)L([\d.]+) ([\d.]+)$/.exec(path.getAttribute("d") ?? "");

                if (segment?.[1] === segment?.[3] && Math.abs(Number(segment?.[4]) - Number(segment?.[2])) > 100) {
                    columns.push(Number(segment?.[1]));
                }

                if (segment?.[2] === segment?.[4] && Number(segment?.[3]) - Number(segment?.[1]) > 100) {
                    rows.push(Number(segment?.[2]));
                }

                // A tick: a short straight segment on either axis.
                if (segment && (segment[1] === segment[3] || segment[2] === segment[4]) && Math.abs(Number(segment[3]) - Number(segment[1]) + Number(segment[4]) - Number(segment[2])) < 10) {
                    ticks++;
                }
            }

            for (const text of renderedPage(presentation).querySelectorAll(".slide-chart text")) {
                labels.push(text.textContent ?? "");
            }

            return { columns: new Set(columns).size, rows: new Set(rows).size, labels, ticks };
        };
        // PptxGenJS 4.0.1 writes majorTickMark and minorTickMark none on both axes, whatever the options say.
        const model = normalizeChart("scatter", xyData, { valAxisMajorTickMark: "outside" });

        expect<unknown[]>([model.categoryMajorTickMark, model.valueAxes[0]?.majorTickMark, model.valueAxes[0]?.minorTickMark]).toEqual(["none", "none", "none"]);

        // LibreOffice runs x 0..7.6 on 0..8 by 1 (nine gridlines) and y 0..6.4 on 0..7 by 1 (eight).
        const automatic = steps({});

        expect([automatic.columns, automatic.rows]).toEqual([9, 8]);
        expect(automatic.labels).toContain("8");
        // valAxisMinVal, valAxisMaxVal and valAxisMajorUnit set the y axis: 0..10 by 2, six gridlines.
        const set = steps({ valAxisMinVal: 0, valAxisMaxVal: 10, valAxisMajorUnit: 2, valAxisLabelFormatCode: "0.0" });

        expect(set.rows).toBe(6);
        expect(set.labels).toContain("10.0");
        expect(set.columns).toBe(9);
        // No ticks on either axis.
        expect([automatic.ticks, set.ticks]).toEqual([0, 0]);
    });

    test("draws a scatter chart's gridlines D9D9D9 on both axes unless valGridLine is set, as PptxGenJS defaults them", () => {
        const xyData: PptxGenJS.OptsChartData[] = [
            { name: "X", values: [1, 2, 3, 4] },
            { name: "Lorem", values: [2, 4, 3, 5] },
        ];

        // PptxGenJS 4.0.1: `catGridLine || (scatter ? { color: 'D9D9D9', size: 1 } : { style: 'none' })`, and valGridLine
        // `|| (scatter ? { color: 'D9D9D9', size: 1 } : {})`, so a bubble chart keeps no x gridlines and 888888 y ones.
        const scatter = normalizeChart("scatter", xyData);

        expect<unknown[]>([scatter.categoryAxis.gridLineColor, scatter.categoryAxis.gridLineStyle, scatter.valueAxes[0]?.gridLineColor, scatter.valueAxes[0]?.gridLineStyle]).toEqual(["#D9D9D9", "solid", "#D9D9D9", "solid"]);
        // A set valGridLine replaces the default whole: a missing color is DEF_CHART_GRIDLINE's 888888.
        expect(normalizeChart("scatter", xyData, { valGridLine: { size: 2 } }).valueAxes[0]).toMatchObject({ gridLineColor: "#888888", gridLineWidth: 2 });
        expect(normalizeChart("scatter", xyData, { valGridLine: { style: "none" } }).valueAxes[0]?.gridLineStyle).toBe("none");

        const bubble = normalizeChart("bubble", [xyData[0]!, { ...xyData[1]!, sizes: [1, 2, 3, 4] }]);

        expect<unknown[]>([bubble.categoryAxis.gridLineStyle, bubble.valueAxes[0]?.gridLineColor]).toEqual(["none", "#888888"]);

        // The vertical (x) and horizontal (y) split lines, by their stroke.
        const gridStrokes = (type: "bubble" | "scatter", options: PptxGenJS.IChartOpts = {}) => {
            const presentation = new PuppeteerGen();
            const data = type === "bubble" ? [xyData[0]!, { ...xyData[1]!, sizes: [1, 2, 3, 4] }] : xyData;

            presentation.addSlide().addChart(type, data, { x: 0, y: 0, w: 5, h: 3, ...options });
            const strokes = { x: new Set<string>(), y: new Set<string>() };

            for (const path of renderedPage(presentation).querySelectorAll('.slide-chart path[fill="none"]')) {
                const segment = path.getAttribute("d")?.match(/^M([\d.]+) ([\d.]+)L([\d.]+) ([\d.]+)$/);
                const stroke = path.getAttribute("stroke");

                if (!segment || !stroke) {
                    continue;
                }

                if (segment[1] === segment[3]) {
                    strokes.x.add(stroke);
                } else if (segment[2] === segment[4]) {
                    strokes.y.add(stroke);
                }
            }

            return { x: [...strokes.x].sort(), y: [...strokes.y].sort() };
        };

        // The 888888 lines are the axis lines.
        expect(gridStrokes("scatter")).toEqual({ x: ["#888888", "#D9D9D9"], y: ["#888888", "#D9D9D9"] });
        expect(gridStrokes("scatter", { valGridLine: { color: "A5A5A5" } })).toEqual({ x: ["#888888", "#D9D9D9"], y: ["#888888", "#A5A5A5"] });
        expect(gridStrokes("bubble")).toEqual({ x: ["#888888"], y: ["#888888"] });
    });

    test("keys bar entries with a square 0.6 of the legend font size and keeps line entries as a line", () => {
        const legendKeys = (type: "bar" | "line", options: PptxGenJS.IChartOpts = {}) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart(type, categoricalData, { x: 0, y: 0, w: 5, h: 3, showLegend: true, legendPos: "b", ...options });

            return Array.from(renderedPage(presentation).querySelectorAll('.slide-chart path[ecmeta_ssr_type="legend"][ecmeta_silent]'), path => path.getAttribute("d"));
        };

        // 10pt legend: 13.33px * 0.6 = 8px squares; 20pt: 16px.
        expect(legendKeys("bar")).toEqual(["M0 0L8 0L8 8L0 8Z", "M0 0L8 0L8 8L0 8Z"]);
        expect(legendKeys("bar", { legendFontSize: 20 })).toEqual(["M0 0L16 0L16 16L0 16Z", "M0 0L16 0L16 16L0 16Z"]);
        expect(legendKeys("line").some(d => d?.includes("L8 8"))).toBe(false);

        const presentation = new PuppeteerGen();

        presentation.addSlide().addChart("bar3D", categoricalData, { x: 0, y: 0, w: 5, h: 3, showLegend: true });
        const key = renderedPage(presentation).querySelector(".slide-chart rect");

        expect([key?.getAttribute("width"), key?.getAttribute("height")]).toEqual(["8", "8"]);
    });

    test("colors each bar of a single-series bar chart from chartColors, as PptxGenJS's per-point c:dPt write it (#473)", () => {
        const one = [categoricalData[0]!];
        const varies = (type: PptxGenJS.CHART_NAME, data: PptxGenJS.OptsChartData[], options: PptxGenJS.IChartOpts) => normalizeChart(type, data, options).series.map(series => series.varyColors);

        expect(varies("bar", one, { chartColors: ["4472C4", "ED7D31"] })).toEqual([true]);
        expect(varies("bar", one, { chartColors: ["4472C4", "ED7D31"], barDir: "bar" })).toEqual([true]);
        // One color, the default palette, lineSize 0 (a c:dPt without a fill), two series or another chart type: none.
        expect(varies("bar", one, { chartColors: ["4472C4"] })).toEqual([false]);
        expect(varies("bar", one, {})).toEqual([false]);
        expect(varies("bar", one, { chartColors: ["4472C4", "ED7D31"], lineSize: 0 })).toEqual([false]);
        expect(varies("bar", categoricalData, { chartColors: ["4472C4", "ED7D31"] })).toEqual([false, false]);
        expect(varies("line", one, { chartColors: ["4472C4", "ED7D31"] })).toEqual([false]);

        const presentation = new PuppeteerGen();

        presentation.addSlide().addChart("bar", one, { x: 0, y: 0, w: 5, h: 3, chartColors: ["4472C4", "ED7D31"], showLegend: true, legendPos: "b" });
        const svg = renderedPage(presentation).querySelector(".slide-chart svg");

        expect(Array.from(svg?.querySelectorAll('path[ecmeta_ssr_type="chart"]') ?? [], path => path.getAttribute("fill"))).toEqual(["#4472C4", "#ED7D31", "#4472C4"]);
        // The legend keys the series in its own color, as LibreOffice draws it.
        expect(Array.from(svg?.querySelectorAll('path[ecmeta_ssr_type="legend"][ecmeta_silent]') ?? [], path => path.getAttribute("fill"))).toEqual(["#4472C4"]);
    });

    test("outlines each box of a single-series bar3D chart from chartColors and fills points without a c:dPt fill as the target does (#507)", () => {
        const one = [categoricalData[0]!];
        const colors = ["FF0000", "00AA00"];
        const points = (type: PptxGenJS.CHART_NAME, options: PptxGenJS.IChartOpts): { fill: string | undefined; lines: string[] | undefined } => {
            const [series] = normalizeChart(type, one, options).series;

            return { fill: series?.automaticPointFill, lines: series?.pointLineColors };
        };

        // PptxGenJS writes each bar3D point an outline in `chartColors[index % length]` and no fill.
        expect(points("bar3D", { chartColors: colors })).toEqual({ fill: "#4472C4", lines: ["#FF0000", "#00AA00", "#FF0000"] });
        // With lineSize 0, a bar3D or bar point has `<a:ln><a:noFill/>` only; one color writes no c:dPt.
        expect(points("bar3D", { chartColors: colors, lineSize: 0 })).toEqual({ fill: "#4472C4", lines: undefined });
        expect(points("bar", { chartColors: colors, lineSize: 0 })).toEqual({ fill: "#4472C4", lines: undefined });
        expect(points("bar3D", { chartColors: ["FF0000"] })).toEqual({ fill: undefined, lines: undefined });
        expect(normalizeChart("bar3D", categoricalData, { chartColors: colors }).series.map(series => series.pointLineColors)).toEqual([undefined, undefined]);

        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("bar3D", one, { x: 0, y: 0, w: 5, h: 3, chartColors: colors, showLegend: true, legendPos: "b" });
        slide.addChart("bar", one, { x: 0, y: 3, w: 5, h: 2, chartColors: colors, lineSize: 0 });
        const [bar3D, bar] = renderedPage(presentation).querySelectorAll(".slide-chart svg");
        const marks = Array.from(bar3D?.querySelectorAll(".bar3d-mark") ?? []);

        // LibreOffice (chartPointKeepsSeriesFill off) fills every box in accent1 (its side 16% darker) and keys the legend
        // in the series color.
        expect(marks.map(mark => mark.getAttribute("stroke"))).toEqual(["#FF0000", "#00AA00", "#FF0000"]);
        expect(marks.map(mark => mark.querySelector(".bar3d-side")?.getAttribute("fill"))).toEqual(["#3960a5", "#3960a5", "#3960a5"]);
        expect(bar3D?.querySelector("rect")?.getAttribute("fill")).toBe("#FF0000");
        expect(Array.from(bar?.querySelectorAll('path[ecmeta_ssr_type="chart"]') ?? [], path => path.getAttribute("fill"))).toEqual(["#4472C4", "#4472C4", "#4472C4"]);
    });

    test("lists each point of a single-series chart with c:dPt in the legend where the target does (chartLegendListsPoints, #546, #552)", () => {
        const one = [categoricalData[0]!];
        const colors = ["FF0000", "00AA00"];
        // Each chart's legendListsPoints, and its legend entries as [text, key fill, key outline].
        const drawn = (target: "libreoffice" | "powerpoint", type: PptxGenJS.CHART_NAME, data: PptxGenJS.OptsChartData[], options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen(undefined, { target });

            presentation.addSlide().addChart(type, data, { x: 0, y: 0, w: 5, h: 3, chartColors: colors, showLegend: true, legendPos: "b", ...options });
            const object = presentation.getModel().slides[0]?.objects[0];
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");
            const entries: (string | null | undefined)[][] = [];

            if (type === "bar3D") {
                for (const text of svg?.querySelectorAll(".bar3d-legend") ?? []) {
                    entries.push([text.textContent, text.previousElementSibling?.getAttribute("fill"), text.previousElementSibling?.getAttribute("stroke")]);
                }
            } else {
                for (const key of svg?.querySelectorAll('path[ecmeta_ssr_type="legend"][ecmeta_silent]') ?? []) {
                    entries.push([key.nextElementSibling?.textContent, key.getAttribute("fill"), key.getAttribute("stroke")]);
                }
            }

            return { listsPoints: object?.kind === "chart" ? object.chart.legendListsPoints : undefined, entries };
        };

        // PowerPoint lists the points of the measured charts, each key drawn as its point: a bar3D box filled in accent1
        // and outlined in its chartColors entry, a lineSize 0 bar in plain accent1.
        expect(drawn("powerpoint", "bar3D", one, {})).toEqual({
            listsPoints: true,
            entries: [
                ["A", "#4472C4", "#FF0000"],
                ["B", "#4472C4", "#00AA00"],
                ["C", "#4472C4", "#FF0000"],
            ],
        });
        expect(drawn("powerpoint", "bar", one, { lineSize: 0 })).toEqual({
            listsPoints: true,
            entries: [
                ["A", "#4472C4", null],
                ["B", "#4472C4", null],
                ["C", "#4472C4", null],
            ],
        });
        // LibreOffice keeps the series entry on the same charts (#545).
        expect(drawn("libreoffice", "bar3D", one, {})).toEqual({ listsPoints: undefined, entries: [["First", "#FF0000", null]] });
        expect(drawn("libreoffice", "bar", one, { lineSize: 0 })).toEqual({ listsPoints: undefined, entries: [["First", "#FF0000", null]] });
        // PowerPoint lists filled c:dPt points too, each keyed in its fill; LibreOffice keeps the series (#552).
        expect(drawn("powerpoint", "bar", one, {})).toEqual({
            listsPoints: true,
            entries: [
                ["A", "#FF0000", null],
                ["B", "#00AA00", null],
                ["C", "#FF0000", null],
            ],
        });
        expect(drawn("libreoffice", "bar", one, {})).toEqual({ listsPoints: undefined, entries: [["First", "#FF0000", null]] });
        // A combo chart's single-series bar type, one chartColors entry (no c:dPt) or two series keep the series entries.
        const combo = new PuppeteerGen(undefined, { target: "powerpoint" });

        (combo.addSlide().addChart as Function)(
            [
                { type: "bar", data: one, options: { chartColors: colors } },
                { type: "line", data: [categoricalData[1]!], options: {} },
            ],
            { x: 0, y: 0, w: 5, h: 3, showLegend: true },
        );
        const comboObject = combo.getModel().slides[0]?.objects[0];

        expect(comboObject?.kind === "chart" ? comboObject.chart.legendListsPoints : "not a chart").toBeUndefined();
        expect(drawn("powerpoint", "bar3D", one, { chartColors: ["FF0000"] })).toEqual({ listsPoints: undefined, entries: [["First", "#FF0000", null]] });
        expect(drawn("powerpoint", "bar", categoricalData, { lineSize: 0 }).listsPoints).toBeUndefined();
    });

    test("runs a clustered horizontal bar chart's series bottom-up and lists its legend from the last series, as LibreOffice and PowerPoint draw it (#472)", () => {
        // For each fill, the top of each bar (by category) and the legend text in order.
        const drawn = (options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("bar", categoricalData, { x: 0, y: 0, w: 5, h: 3, chartColors: ["4472C4", "ED7D31"], showLegend: true, legendPos: "b", ...options });
            const svg = renderedPage(presentation).querySelector(".slide-chart svg");
            const tops: Record<string, number[]> = {};

            for (const path of svg?.querySelectorAll('path[ecmeta_ssr_type="chart"]') ?? []) {
                const fill = path.getAttribute("fill") ?? "";

                (tops[fill] ??= []).push(Number(/^M[\d.]+ ([\d.]+)/.exec(path.getAttribute("d") ?? "")?.[1]));
            }

            // The legend's texts are the only ones that name a series.
            const legend: string[] = [];

            for (const text of svg?.querySelectorAll("text") ?? []) {
                if (text.textContent === "First" || text.textContent === "Second") {
                    legend.push(text.textContent);
                }
            }

            return { tops, legend };
        };
        const clustered = drawn({ barDir: "bar" });
        const blue = clustered.tops["#4472C4"] ?? [];
        const orange = clustered.tops["#ED7D31"] ?? [];

        expect(blue).toHaveLength(3);
        expect(orange).toHaveLength(3);

        // The first series ("First", blue) sits below the second in every band.
        for (const [index, top] of blue.entries()) {
            expect(top).toBeGreaterThan(orange[index] ?? Infinity);
        }

        expect(clustered.legend).toEqual(["Second", "First"]);
        // A stacked one keeps the series order in the legend; a column chart too.
        expect(drawn({ barDir: "bar", barGrouping: "stacked" }).legend).toEqual(["First", "Second"]);
        expect(drawn({ barDir: "col" }).legend).toEqual(["First", "Second"]);
    });

    test("lists a stacked column or area chart's legend from the last series when it is vertical, as LibreOffice draws it (#508)", () => {
        const legend = (type: "area" | "bar", options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart(type, categoricalData, { x: 0, y: 0, w: 5, h: 3, showLegend: true, ...options });
            const names: string[] = [];

            for (const text of renderedPage(presentation).querySelectorAll(".slide-chart svg text")) {
                if (text.textContent === "First" || text.textContent === "Second") {
                    names.push(text.textContent);
                }
            }

            return names;
        };

        // LibreOffice 26.2: the stack's order, bottom-up, in a right, left or top-right legend.
        expect(legend("bar", { barGrouping: "stacked", legendPos: "r" })).toEqual(["Second", "First"]);
        expect(legend("bar", { barGrouping: "stacked", legendPos: "l" })).toEqual(["Second", "First"]);
        expect(legend("bar", { barGrouping: "percentStacked", legendPos: "tr" })).toEqual(["Second", "First"]);
        expect(legend("area", { barGrouping: "stacked", legendPos: "r" })).toEqual(["Second", "First"]);
        // Series order in a top or bottom legend, beside clustered columns and standard areas, and beside a stacked
        // horizontal bar chart, whose stack runs left to right.
        expect(legend("bar", { barGrouping: "stacked", legendPos: "b" })).toEqual(["First", "Second"]);
        expect(legend("bar", { barGrouping: "stacked", legendPos: "t" })).toEqual(["First", "Second"]);
        expect(legend("bar", { legendPos: "r" })).toEqual(["First", "Second"]);
        expect(legend("area", { legendPos: "r" })).toEqual(["First", "Second"]);
        expect(legend("bar", { barDir: "bar", barGrouping: "stacked", legendPos: "r" })).toEqual(["First", "Second"]);
    });

    test("gives a stacked bar chart PptxGenJS's 50% gap width when barGapWidthPct is unset or 0", () => {
        // PptxGenJS 4.0.1: `if (barGrouping.includes('tacked') && !barGapWidthPct) barGapWidthPct = 50`, then `|| 150`.
        expect(normalizeChart("bar", categoricalData, { barGrouping: "stacked" }).barGapWidthPercent).toBe(50);
        expect(normalizeChart("bar", categoricalData, { barGrouping: "percentStacked", barGapWidthPct: 0 }).barGapWidthPercent).toBe(50);
        expect(normalizeChart("bar3D", categoricalData, { barGrouping: "stacked" }).barGapWidthPercent).toBe(50);
        expect(normalizeChart("bar", categoricalData, { barGrouping: "stacked", barGapWidthPct: 300 }).barGapWidthPercent).toBe(300);
        expect(normalizeChart("bar", categoricalData).barGapWidthPercent).toBe(150);
        expect(normalizeChart("bar", categoricalData, { barGapWidthPct: 0 }).barGapWidthPercent).toBe(0);

        // Each first-series bar's thickness (a horizontal bar's height) in px.
        const thickness = (options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("bar", categoricalData, { x: 0, y: 0, w: 5, h: 3, barDir: "bar", chartColors: ["4472C4", "ED7D31"], ...options });
            const heights = new Set<number>();

            for (const path of renderedPage(presentation).querySelectorAll('.slide-chart path[ecmeta_ssr_type="chart"][fill="#4472C4"]')) {
                // ECharts draws a bar as `M x y l w 0 l 0 h l -w 0 Z`.
                heights.add(Number(/l0 ([\d.]+)/.exec(path.getAttribute("d") ?? "")?.[1]));
            }

            return [...heights];
        };
        const [stacked] = thickness({ barGrouping: "stacked" });
        const [clustered] = thickness({});

        // In the same band, a stacked bar is band / 1.5 and each of two clustered bars band / (2 + 1.5).
        expect(stacked).toBeDefined();
        expect((stacked ?? 0) / (clustered ?? 1)).toBeCloseTo(3.5 / 1.5, 1);
    });

    test("draws outside major tick marks on the category axis, as PptxGenJS writes them", () => {
        // PptxGenJS 4.0.1: `<c:majorTickMark val="${catAxisMajorTickMark || 'out'}"/>` on every category axis but scatter's.
        expect(normalizeChart("bar", categoricalData).categoryMajorTickMark).toBe("outside");
        const presentation = new PuppeteerGen();

        presentation.addSlide().addChart("bar", categoricalData, { x: 0, y: 0, w: 5, h: 3, valAxisMajorTickMark: "none" });
        const paths = Array.from(renderedPage(presentation).querySelectorAll(".slide-chart path"), path => path.getAttribute("d") ?? "");
        // A vertical 5px stroke at each of the four category boundaries of three categories.
        const ticks = paths.filter(d => {
            const match = d.match(/^M([\d.]+) ([\d.]+)L([\d.]+) ([\d.]+)$/);

            return match !== null && match[1] === match[3] && Math.abs(Number(match[4]) - Number(match[2]) - 5) < 0.01;
        });

        expect(ticks).toHaveLength(4);
    });

    test("draws the bar3D legend in the legend font, not the axis-label font", () => {
        const legendText = (options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("bar3D", categoricalData, { x: 0, y: 0, w: 5, h: 3, showLegend: true, ...options });
            const text = renderedPage(presentation).querySelector(".bar3d-legend");

            return [text?.getAttribute("font-family"), text?.getAttribute("font-size"), text?.getAttribute("fill")];
        };

        // PptxGenJS writes no legend <c:txPr> by default: Office's 10pt chart text in the body font.
        expect(legendText({})).toEqual(["Calibri", "13.33", "#000000"]);
        expect(legendText({ legendFontFace: "Georgia", legendFontSize: 14, legendColor: "C00000" })).toEqual(["Georgia", "18.67", "#C00000"]);
    });

    test("normalizes and renders primary-axis bar, line, and area combinations", () => {
        const mixed: PptxGenJS.IChartMulti[] = [
            { type: "bar", data: [categoricalData[0]!], options: {} },
            { type: "line", data: [categoricalData[1]!], options: {} },
            { type: "area", data: [{ name: "Third", labels: ["A", "B", "C"], values: [1, 3, 2] }], options: {} },
        ];
        const chart = normalizeChart(mixed, [], { chartColors: ["4472C4", "ED7D31", "70AD47"] });

        expect(chart.type).toBe("mixed");
        expect(chart.series.map(series => series.type)).toEqual(["bar", "line", "area"]);

        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        expect(
            (slide.addChart as Function)(mixed, {
                x: 1,
                y: 1,
                w: 6,
                h: 3,
                showLegend: true,
                chartColors: ["4472C4", "ED7D31", "70AD47"],
            }),
        ).toBe(slide);
        const svg = renderedPage(presentation).querySelector(".slide-chart svg");

        expect(svg?.textContent).toContain("First");
        expect(svg?.textContent).toContain("Second");
        expect(svg?.textContent).toContain("Third");
    });

    test("normalizes and renders a secondary value axis for mixed charts", () => {
        const mixed: PptxGenJS.IChartMulti[] = [
            { type: "bar", data: [categoricalData[0]!], options: {} },
            {
                type: "line",
                data: [{ name: "Revenue", labels: ["A", "B", "C"], values: [100, 400, 250] }],
                options: { secondaryValAxis: true },
            },
        ];
        const options: PptxGenJS.IChartOpts = {
            valAxes: [
                { valAxisMinVal: 0, valAxisMaxVal: 8, showValAxisTitle: true, valAxisTitle: "Units" },
                { valAxisMinVal: 0, valAxisMaxVal: 500, showValAxisTitle: true, valAxisTitle: "Revenue" },
            ],
        };
        const chart = normalizeChart(mixed, [], options);

        expect(chart.series.map(series => series.valueAxisIndex)).toEqual([0, 1]);
        expect(chart.valueAxes).toMatchObject([
            { minimum: 0, maximum: 8, hidden: false, showTitle: true, title: "Units" },
            { minimum: 0, maximum: 500, hidden: false, showTitle: true, title: "Revenue" },
        ]);

        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        (slide.addChart as Function)(mixed, { ...options, x: 1, y: 1, w: 6, h: 3 });
        const svg = renderedPage(presentation).querySelector(".slide-chart svg");

        expect(svg?.textContent).toContain("Units");
        expect(svg?.textContent).toContain("Revenue");
        expect(svg?.textContent).toContain("500");
    });

    test("normalizes advanced value-axis formatting and renders formatted ticks", () => {
        const options: PptxGenJS.IChartOpts = {
            valAxisMinVal: 0,
            valAxisMaxVal: 8,
            valAxisMajorUnit: 2,
            valAxisLabelColor: "C00000",
            valAxisLabelFontBold: true,
            valAxisLabelFontFace: "Arial",
            valAxisLabelFontItalic: true,
            valAxisLabelFontSize: 14,
            valAxisLabelFormatCode: "$0.0",
            valAxisLabelPos: "high",
            valAxisLabelRotate: 15,
            valAxisLineColor: "00AA00",
            valAxisLineShow: true,
            valAxisLineSize: 2,
            valAxisLineStyle: "dash",
            valAxisMajorTickMark: "inside",
            valAxisMinorTickMark: "outside",
            valAxisOrientation: "minMax",
            showValAxisTitle: true,
            valAxisTitle: "Revenue",
            valAxisTitleColor: "0000FF",
            valAxisTitleFontFace: "Georgia",
            valAxisTitleFontSize: 16,
            valAxisTitleRotate: 90,
            valGridLine: { color: "FF00FF", size: 1.5, style: "dot", cap: "round" },
        };
        const axis = normalizeChart("bar", categoricalData, options).valueAxes[0]!;

        expect(axis).toMatchObject({
            minimum: 0,
            maximum: 8,
            majorUnit: 2,
            labelColor: "#C00000",
            labelBold: true,
            labelItalic: true,
            labelFontFace: "Arial",
            labelFontSize: 14,
            labelFormatCode: "$0.0",
            labelPosition: "high",
            labelRotate: 15,
            lineColor: "#00AA00",
            lineWidth: 2,
            lineStyle: "dash",
            lineVisible: true,
            majorTickMark: "inside",
            minorTickMark: "outside",
            title: "Revenue",
            titleColor: "#0000FF",
            titleFontFace: "Georgia",
            titleFontSize: 16,
            titleRotate: 90,
            gridLineColor: "#FF00FF",
            gridLineWidth: 1.5,
            gridLineStyle: "dot",
            gridLineCap: "round",
        });

        const presentation = new PuppeteerGen();

        presentation.addSlide().addChart("bar", categoricalData, { ...options, x: 1, y: 1, w: 6, h: 3 });
        const svg = renderedPage(presentation).querySelector(".slide-chart svg");

        expect(svg?.textContent).toContain("Revenue");
        expect(svg?.textContent).toContain("$8.0");
        expect(formatChartNumber(1234.5, "$#,##0.00")).toBe("$1,234.50");
        expect(formatChartNumber(0.25, "0%")).toBe("25%");
    });

    test("validates logarithmic axes and explicitly rejects unsupported formats", () => {
        const logarithmic = normalizeChart("line", [{ name: "Growth", labels: ["A", "B"], values: [1, 100] }], {
            valAxisMinVal: 1,
            valAxisMaxVal: 100,
            valAxisLogScaleBase: 10,
        });

        expect(logarithmic.valueAxes[0]?.logScaleBase).toBe(10);
        expect(() => normalizeChart("line", categoricalData, { valAxisLogScaleBase: 10 })).toThrow("positive minimum");
        expect(() => normalizeChart("line", categoricalData, { valAxisLogScaleBase: 1 })).toThrow("between 2 and 99");
        expect(() => normalizeChart("line", categoricalData, { valAxisLabelFormatCode: "0;[Red]-0" })).toThrow(UnsupportedChartError);
        // LibreOffice 26.2.6 draws a cross tick across the axis line (#582).
        expect(normalizeChart("line", categoricalData, { valAxisMajorTickMark: "cross", catAxisMinorTickMark: "cross" }).valueAxes[0]?.majorTickMark).toBe("cross");
    });

    test("rejects incompatible mixed families and unimplemented per-series axes", () => {
        expect(() =>
            normalizeChart(
                [
                    { type: "bar", data: [categoricalData[0]!], options: {} },
                    { type: "pie", data: [categoricalData[1]!], options: {} },
                ],
                [],
            ),
        ).toThrow(UnsupportedChartError);
        try {
            normalizeChart(
                [
                    { type: "bar", data: [categoricalData[0]!], options: {} },
                    { type: "line", data: [categoricalData[1]!], options: { secondaryCatAxis: true } },
                ],
                [],
            );
            throw new Error("Expected normalization to fail");
        } catch (error) {
            expect(error).toBeInstanceOf(UnsupportedChartError);
            expect((error as UnsupportedChartError).reason).toBe("chart-option");
            expect((error as UnsupportedChartError).unsupportedOptions).toEqual(["series[1].options.secondaryCatAxis[needs secondaryValAxis and catAxes[1]]"]);
        }

        expect(() =>
            normalizeChart(
                [
                    { type: "bar", data: [categoricalData[0]!], options: {} },
                    { type: "line", data: [categoricalData[1]!], options: { secondaryValAxis: true } },
                ],
                [],
            ),
        ).toThrow("valAxes[primary,secondary]");
    });

    test("supports native ECharts and trusted SVG extension inputs", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        expect(
            slide.addChartEx(
                {
                    renderer: "echarts",
                    option: { xAxis: { data: ["A"] }, yAxis: {}, series: [{ type: "bar", data: [4] }] },
                },
                { x: 0, y: 0, w: 3, h: 2, objectName: "Native" },
            ),
        ).toBe(slide);
        slide.addChartEx(
            {
                renderer: "svg",
                svg: '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><rect id="face" width="10" height="10"/></svg>',
            },
            { x: 3, y: 0, w: 1, h: 1, objectName: "Raw SVG" },
        );

        expect(renderedPage(presentation).querySelectorAll(".slide-chart svg")).toHaveLength(2);
        expect(renderedPage(presentation).querySelector("#slide-1-object-1-face")).not.toBeNull();
    });

    test("rejects external resources in extension SVG", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        expect(() =>
            slide.addChartEx(
                {
                    renderer: "svg",
                    svg: '<svg xmlns="http://www.w3.org/2000/svg"><image href="https://example.com/a.png"/></svg>',
                },
                { x: 0, y: 0, w: 1, h: 1 },
            ),
        ).toThrow("must not reference external resources");
    });
});

describe("chart axis and plot-area options on every chart type (#561, #479)", () => {
    const quarters = [{ name: "North", labels: ["Q1", "Q2", "Q3", "Q4"], values: [3, 5, 4, 7] }];
    const chartSvgs = (presentation: PuppeteerGen) => Array.from(renderedPage(presentation).querySelectorAll(".slide-chart svg"));
    const textsOf = (svg: Element | undefined) => Array.from(svg?.querySelectorAll("text") ?? [], text => text.textContent ?? "");
    // ECharts places its text by a translate transform.
    const textY = (svg: Element | undefined, content: string) =>
        Number(
            Array.from(svg?.querySelectorAll("text") ?? [])
                .find(text => text.textContent === content)
                ?.getAttribute("transform")
                ?.match(/translate\([\d.]+ ([\d.]+)/)?.[1] ?? Number.NaN,
        );

    test("scales a radar value axis to at most two steps, as LibreOffice does", () => {
        const scale = (values: number[], options: PptxGenJS.IChartOpts = {}) => {
            const axis = normalizeChart("radar", [{ name: "North", labels: ["A", "B", "C", "D", "E"], values }], options).valueAxes[0];

            return [axis?.minimum, axis?.maximum, axis?.majorUnit];
        };

        // LibreOffice 26.2.6 renders of these PptxGenJS 4.0.1 radar charts at any size or label font.
        expect(scale([3, 5, 4, 7, 2])).toEqual([0, 10, 5]);
        expect(scale([13, 5, 23, 17, 2])).toEqual([0, 40, 20]);
        expect(scale([30, 50, 40, 70, 20])).toEqual([0, 100, 50]);
        expect(scale([0.3, 0.5, 0.4, 0.8, 0.2])).toEqual([0, 1, 0.5]);
        expect(scale([3, -5, 4, 7, 2])).toEqual([-10, 10, 10]);
        expect(scale([3, 5, 4, 7, 2], { valAxisMaxVal: 8, valAxisMajorUnit: 2 })).toEqual([0, 8, 2]);
    });

    test("scales a radar value axis as a 2D axis in at most seven steps on the PowerPoint target (radarScale, #565)", () => {
        const scale = (values: number[]) => {
            const axis = normalizeChart("radar", [{ name: "North", labels: ["A", "B", "C", "D", "E"], values }], {}, "powerpoint").valueAxes[0];

            return [axis?.minimum, axis?.maximum, axis?.majorUnit];
        };

        // PowerPoint's PDF export: 0..8 by 2 for data up to 7 (#568), 0..25 by 5 up to 23 and 0..140 by 20 up to 130.
        expect(scale([3, 5, 4, 7, 2])).toEqual([0, 8, 2]);
        expect(scale([12, 23, 8, 17, 20])).toEqual([0, 25, 5]);
        expect(scale([40, 130, 75, 110, 90])).toEqual([0, 140, 20]);
    });

    test("draws a radar chart's value labels, rings, spokes, category labels and series styles", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("radar", quarters, { x: 0, y: 0, w: 4, h: 3, valAxisLabelFormatCode: "0.0", valGridLine: { color: "FF0000" }, valAxisLineColor: "0000FF", catAxisLabelColor: "00B050", lineSize: 4 });
        slide.addChart("radar", quarters, { x: 4, y: 0, w: 4, h: 3, valAxisHidden: true, radarStyle: "filled", chartColorsOpacity: 40 });
        slide.addChart("radar", quarters, { x: 0, y: 3, w: 4, h: 2, layout: { x: 0, y: 0, w: 0.5, h: 1 } });
        const [styled, hidden, placed] = chartSvgs(presentation);

        // Only the first spoke is numbered.
        expect(textsOf(styled).filter(text => /^\d/.test(text))).toEqual(["0.0", "5.0", "10.0"]);
        expect(styled?.querySelector('[stroke="#FF0000"]')).not.toBeNull();
        expect(styled?.querySelector('[stroke="#0000FF"]')).not.toBeNull();
        expect(styled?.querySelector('text[fill="#00B050"]')?.textContent).toBe("Q1");
        expect(styled?.querySelector('[stroke-width^="5.33"]')).not.toBeNull();
        expect(textsOf(hidden)).toEqual(["Q1", "Q2", "Q3", "Q4"]);
        expect(hidden?.querySelector('[fill-opacity="0.4"]')).not.toBeNull();
        // Office runs the categories clockwise from the top: Q2 is right of the center.
        const q2 = Array.from(styled?.querySelectorAll("text") ?? []).find(text => text.textContent === "Q2");

        expect(Number(q2?.getAttribute("transform")?.match(/translate\(([\d.]+)/)?.[1])).toBeGreaterThan(inchesToPixels(4) / 2);
        // The layout box's center.
        const q1 = Array.from(placed?.querySelectorAll("text") ?? []).find(text => text.textContent === "Q1");

        expect(Number(q1?.getAttribute("transform")?.match(/translate\(([\d.]+)/)?.[1])).toBeCloseTo(inchesToPixels(4) / 4, 0);
    });

    test("rejects radar value and category options LibreOffice does not draw, naming them", () => {
        expect(() => normalizeChart("radar", quarters, { valAxisLabelRotate: 45 })).toThrow("valAxisLabelRotate[radar]");
        expect(() => normalizeChart("radar", quarters, { valAxisLogScaleBase: 10 })).toThrow("valAxisLogScaleBase[radar]");
        expect(() => normalizeChart("radar", quarters, { valAxisOrientation: "maxMin" } as unknown as PptxGenJS.IChartOpts)).toThrow("valAxisOrientation[radar]");
        expect(() => normalizeChart("radar", quarters, { valAxisLabelPos: "high" })).toThrow("valAxisLabelPos[high][radar]");
        expect(() => normalizeChart("radar", quarters, { valAxisLabelPos: "nextTo", valAxisLabelRotate: 0, valAxisOrientation: "minMax", showValAxisTitle: true })).not.toThrow();
    });

    test("draws a bar3D chart's value axis options and rejects the ones it cannot draw", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("bar3D", quarters, { x: 0, y: 0, w: 4, h: 3, valAxisHidden: true });
        slide.addChart("bar3D", quarters, { x: 4, y: 0, w: 4, h: 3, showValAxisTitle: true, valAxisTitle: "Units", valAxisLabelRotate: 45, valGridLine: { color: "FF0000", size: 2 } });
        slide.addChart("bar3D", quarters, { x: 0, y: 3, w: 4, h: 2.5, valGridLine: { style: "none" }, valAxisLabelPos: "high" });
        const [hidden, titled, high] = chartSvgs(presentation);

        expect(textsOf(hidden)).toEqual(["Q1", "Q2", "Q3", "Q4"]);
        expect(titled?.querySelector(".bar3d-value-title")?.textContent).toBe("Units");
        // The title keeps a band of its own inside the chart, left of the value labels (#568).
        const titleX = Number(titled?.querySelector(".bar3d-value-title")?.getAttribute("x"));
        const labelXs = Array.from(titled?.querySelectorAll('text[text-anchor="end"]') ?? [], text => Number(text.getAttribute("x")));

        expect(titleX).toBeGreaterThan(0);
        expect(titleX).toBeLessThan(Math.min(...labelXs));
        expect(titled?.querySelector('text[transform^="rotate(45 "]')).not.toBeNull();
        expect(titled?.querySelectorAll('polyline[stroke="#FF0000"]').length).toBeGreaterThan(1);
        // No gridlines, and the labels at the far edge of the back wall (right of the category labels).
        expect(high?.querySelectorAll('polyline:not(.bar3d-frame)[stroke="#888888"]')).toHaveLength(0);
        const zero = Array.from(high?.querySelectorAll("text") ?? []).find(text => text.textContent === "0");

        expect(Number(zero?.getAttribute("x"))).toBeGreaterThan(inchesToPixels(4) / 2);
        // LibreOffice draws no axis line or tick marks on a 3D wall (drawsBar3DValueAxisLine off): accepted, drawn as none.
        expect(normalizeChart("bar3D", quarters, { valAxisLineColor: "FF0000", valAxisLineSize: 3, valAxisMinorTickMark: "outside" }).valueAxes[0]).toMatchObject({ lineVisible: false, majorTickMark: "none", minorTickMark: "none" });
        expect(() => normalizeChart("bar3D", quarters, { valAxisLogScaleBase: 10 })).toThrow("valAxisLogScaleBase[bar3D]");
        expect(() => normalizeChart("bar3D", quarters, { valAxisOrientation: "maxMin" } as unknown as PptxGenJS.IChartOpts)).toThrow("valAxisOrientation[bar3D]");
        expect(() => normalizeChart("bar3D", quarters, { barDir: "bar" })).toThrow("barDir[bar3D]");
    });

    test("draws a bar3D chart's category labels, title and gridlines, its legend position and its layout", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("bar3D", quarters, { x: 0, y: 0, w: 4, h: 3, catAxisLabelColor: "C00000", catAxisLabelFontBold: true, showCatAxisTitle: true, catAxisTitle: "Quarter", catGridLine: { color: "00B050" } });
        slide.addChart("bar3D", quarters, { x: 4, y: 0, w: 4, h: 3, showLegend: true, legendPos: "r" });
        slide.addChart("bar3D", quarters, { x: 0, y: 3, w: 4, h: 2.5, layout: { x: 0, y: 0, w: 0.5, h: 0.5 } });
        const [category, legend, placed] = chartSvgs(presentation);
        const q1 = category?.querySelector('text[fill="#C00000"]');

        expect(q1?.textContent).toBe("Q1");
        expect(q1?.getAttribute("font-weight")).toBe("bold");
        expect(category?.querySelector(".bar3d-category-title")?.textContent).toBe("Quarter");
        // A line at each of the five category boundaries.
        expect(category?.querySelectorAll(".bar3d-category-grid")).toHaveLength(5);
        expect(Number(legend?.querySelector(".bar3d-legend")?.getAttribute("x"))).toBeGreaterThan(inchesToPixels(4) * 0.75);
        // Every category label within the layout's left half.
        expect(
            Math.max(
                ...["Q1", "Q2", "Q3", "Q4"].map(label =>
                    Number(
                        Array.from(placed?.querySelectorAll("text") ?? [])
                            .find(text => text.textContent === label)
                            ?.getAttribute("x"),
                    ),
                ),
            ),
        ).toBeLessThan(inchesToPixels(4) / 2);
    });

    test("reverses a value axis for valAxisOrientation maxMin and validates the value", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("line", quarters, { x: 0, y: 0, w: 4, h: 3, valAxisOrientation: "maxMin" } as unknown as PptxGenJS.IChartOpts);
        slide.addChart("line", quarters, { x: 4, y: 0, w: 4, h: 3 });
        const [reversed, plain] = chartSvgs(presentation);

        expect(textY(reversed, "0")).toBeLessThan(textY(reversed, "8"));
        expect(textY(reversed, "Q1")).toBeLessThan(textY(reversed, "0"));
        expect(textY(plain, "0")).toBeGreaterThan(textY(plain, "8"));
        expect(() => normalizeChart("line", quarters, { valAxisOrientation: "sideways" } as unknown as PptxGenJS.IChartOpts)).toThrow("valAxisOrientation");
    });

    test("never smooths an area chart, as PptxGenJS writes <c:smooth> on line series only", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("area", quarters, { x: 0, y: 0, w: 4, h: 3, lineSmooth: true });
        slide.addChart("line", quarters, { x: 4, y: 0, w: 4, h: 3, lineSmooth: true });
        const [area, line] = chartSvgs(presentation);
        const curved = (svg: Element | undefined) => Array.from(svg?.querySelectorAll("path") ?? []).some(path => /C/.test(path.getAttribute("d") ?? ""));

        expect(curved(area)).toBe(false);
        expect(curved(line)).toBe(true);
    });

    test("draws a scatter chart's value axis options and its x axis in the category axis style", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("scatter", scatterData, { x: 0, y: 0, w: 4, h: 3, valAxisHidden: true, showCatAxisTitle: true, catAxisTitle: "Width", catAxisLabelColor: "C00000" });
        slide.addChart("scatter", scatterData, { x: 4, y: 0, w: 4, h: 3, showValAxisTitle: true, valAxisTitle: "Height", valAxisLabelFormatCode: "0.0" });
        const [hidden, titled] = chartSvgs(presentation);

        expect(textsOf(hidden)).toContain("Width");
        expect(textsOf(hidden).some(text => text === "6" || text === "7")).toBe(false);
        expect(hidden?.querySelector('text[fill="#C00000"]')).not.toBeNull();
        expect(textsOf(titled)).toEqual(expect.arrayContaining(["Height", "6.0"]));
    });

    test("places a pie chart by its manual layout", () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addChart("pie", quarters, { x: 0, y: 0, w: 4, h: 3, layout: { x: 0, y: 0, w: 0.5, h: 0.5 } });
        const [pie] = chartSvgs(presentation);
        let right = 0;

        for (const path of Array.from(pie?.querySelectorAll("path") ?? [])) {
            for (const match of (path.getAttribute("d") ?? "").matchAll(/(?:^|[A-Za-z ])([\d.]+) [\d.]+/g)) {
                right = Math.max(right, Number(match[1]));
            }
        }

        expect(right).toBeLessThanOrEqual(inchesToPixels(4) / 2 + 1);
    });
});

describe("chart category-axis, data-table and point options (#566, #567, #548, #537)", () => {
    const quarters = [{ name: "Lorem", labels: ["Q1", "Q2", "Q3", "Q4"], values: [3, -2, 4, 7] }];
    const pair = [
        { name: "Lorem", labels: ["Q1", "Q2", "Q3", "Q4"], values: [3, 5, 4, 7] },
        { name: "Ipsum", labels: ["Q1", "Q2", "Q3", "Q4"], values: [5, 2, 6, 4] },
    ];
    const chartSvg = (type: PptxGenJS.CHART_NAME, data: PptxGenJS.OptsChartData[], options: PptxGenJS.IChartOpts) => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addChart(type, data, { x: 0, y: 0, w: 6, h: 4, ...options });

        return renderedPage(presentation).querySelector(".slide-chart svg");
    };
    const textsOf = (svg: Element | null) => Array.from(svg?.querySelectorAll("text") ?? [], text => text.textContent ?? "");

    test("normalizes the category-axis options from PptxGenJS's defaults", () => {
        const axis = normalizeChart("bar", quarters, {
            catAxisHidden: true,
            catAxisLabelPos: "high",
            catAxisLabelRotate: 45,
            catAxisLineColor: "FF0000",
            catAxisLineSize: 3,
            catAxisLineStyle: "dash",
            catAxisMinorTickMark: "outside",
            catAxisOrientation: "maxMin",
            catAxisCrossesAt: 2,
        } as unknown as PptxGenJS.IChartOpts).categoryAxis;

        expect(axis).toMatchObject({ hidden: true, labelPosition: "high", labelRotate: 45, lineColor: "#FF0000", lineWidth: 3, lineStyle: "dash", minorTickMark: "outside", reversed: true, valueAxisCrossing: { kind: "category", index: 1 } });
        // PptxGenJS writes `low` for a column chart's labels and `nextTo` for a bar chart's.
        expect(normalizeChart("bar", quarters).categoryAxis.labelPosition).toBe("low");
        expect(normalizeChart("bar", quarters, { barDir: "bar" }).categoryAxis.labelPosition).toBe("nextTo");
        expect(normalizeChart("bar", quarters, { catAxisCrossesAt: "max" } as unknown as PptxGenJS.IChartOpts).categoryAxis.valueAxisCrossing).toEqual({ kind: "max" });
        expect(normalizeChart("bar", quarters, { valAxisCrossesAt: 4 } as PptxGenJS.IChartOpts).valueAxes[0]?.categoryAxisCrossing).toEqual({ kind: "value", value: 4 });
        expect(normalizeChart("bar", [{ name: "Lorem", labels: ["1", "2"], values: [3, 5] }], { catLabelFormatCode: "0.0" }).categoryAxis.labelFormatCode).toBe("0.0");
    });

    test("reads label frequency and display units only where the target does", () => {
        const options = { catAxisLabelFrequency: 2, valAxisDisplayUnit: "thousands", valAxisDisplayUnitLabel: true } as unknown as PptxGenJS.IChartOpts;
        // LibreOffice 26.2 ignores `<c:tickLblSkip>` and `<c:dispUnits>` (readsCategoryLabelSkip, readsValueAxisDisplayUnits).
        const libreOffice = normalizeChart("bar", quarters, options, "libreoffice");
        const powerPoint = normalizeChart("bar", quarters, options, "powerpoint");

        expect(libreOffice.categoryAxis.labelSkip).toBeUndefined();
        expect(libreOffice.valueAxes[0]?.displayUnit).toBeUndefined();
        expect(powerPoint.categoryAxis.labelSkip).toBe(2);
        expect(powerPoint.valueAxes[0]?.displayUnit).toEqual({ divisor: 1000, label: PPTX_DEFAULTS.chart.displayUnits.thousands.label });
        expect(() => normalizeChart("bar", quarters, { valAxisDisplayUnit: "lorem" } as unknown as PptxGenJS.IChartOpts)).toThrow("valAxisDisplayUnit");
        expect(() => normalizeChart("bar", quarters, { ...options, showValAxisTitle: true }, "powerpoint")).toThrow("valAxisDisplayUnitLabel[showValAxisTitle]");
    });

    test("rejects the category-axis options a chart type cannot draw, naming them", () => {
        expect(() => normalizeChart("radar", quarters, { catAxisLabelRotate: 45 })).toThrow("catAxisLabelRotate[radar]");
        expect(() => normalizeChart("bar3D", quarters, { catAxisBaseTimeUnit: "days" } as PptxGenJS.IChartOpts)).toThrow("catAxisBaseTimeUnit[bar3D]");
        // A scatter chart's x axis is a value axis: the y axis crosses it at that x value (#591).
        const xy = [
            { name: "X", values: [1, 2, 3, 4, 5, 6] },
            { name: "Lorem", values: [10, 22, 15, 30, 25, 18] },
        ];
        const crossed = normalizeChart("scatter", xy, { catAxisCrossesAt: 3, valAxisCrossesAt: 20 } as PptxGenJS.IChartOpts);

        expect(crossed.categoryAxis.valueAxisCrossing).toEqual({ kind: "value", value: 3 });
        expect(crossed.valueAxes[0]?.categoryAxisCrossing).toEqual({ kind: "value", value: 20 });
        // PowerPoint's PDF export draws a bubble chart's y axis at x 0.5, 1.5, 2.5 and 19.5 for catAxisCrossesAt 1, 2, 3
        // and 20 (bubbleCategoryCrossingShift), LibreOffice at that x value.
        const bubbleCrossing = (target: "libreoffice" | "powerpoint", type: "bubble" | "scatter"): unknown => {
            const presentation = new PuppeteerGen(undefined, { target });
            const data = type === "bubble" ? [xy[0], { ...xy[1], sizes: [4, 9, 6, 12, 5, 7] }] : xy;

            presentation.addSlide().addChart(type, data, { x: 1, y: 1, w: 4, h: 3, catAxisCrossesAt: 2 } as PptxGenJS.IChartOpts);
            const chart = presentation.getModel().slides[0]?.objects[0];

            return chart?.kind === "chart" ? chart.chart.categoryAxis.valueAxisCrossing : undefined;
        };

        expect(bubbleCrossing("powerpoint", "bubble")).toEqual({ kind: "value", value: 1.5 });
        expect(bubbleCrossing("powerpoint", "scatter")).toEqual({ kind: "value", value: 2 });
        expect(bubbleCrossing("libreoffice", "bubble")).toEqual({ kind: "value", value: 2 });
        expect(() => normalizeChart("scatter", xy, { catAxisCrossesAt: Number.NaN } as PptxGenJS.IChartOpts)).toThrow("catAxisCrossesAt");
        expect(() => normalizeChart("bar", quarters, { catAxisMajorTickMark: "in" } as unknown as PptxGenJS.IChartOpts)).toThrow(TypeError);
        // A date format needs date serials: text categories are not a time line.
        expect(() => normalizeChart("line", quarters, { catLabelFormatCode: "mmm-yy" })).toThrow("catLabelFormatCode[format]");
        expect(() => normalizeChart("line", [{ name: "Lorem", labels: ["45292"], values: [1] }], { catLabelFormatCode: "h:mm" })).toThrow("catLabelFormatCode[format]");
        expect(() => normalizeChart("bar3D", quarters, { dataLabelBkgrdColors: true })).toThrow("dataLabelBkgrdColors[bar3D]");
        expect(() => normalizeChart("bar", quarters, { catAxisOrientation: "minMax", valAxisCrossesAt: "autoZero" } as PptxGenJS.IChartOpts)).not.toThrow();
    });

    test("draws a bar3D chart's axis crossings, high category labels and data table as both targets do (#591, #592)", () => {
        const bar3D = (options: PptxGenJS.IChartOpts) => chartSvg("bar3D", pair, options);
        const x = (text: Element | undefined) => Number(text?.getAttribute("x"));
        const y = (text: Element | undefined) => Number(text?.getAttribute("y"));
        const textNamed = (svg: Element | null, content: string) => Array.from(svg?.querySelectorAll("text") ?? []).find(text => text.textContent === content);
        const fronts = (svg: Element | null) => Array.from(svg?.querySelectorAll(".bar3d-front") ?? []);
        const plain = bar3D({});

        // catAxisCrossesAt 2 puts the value labels between the first and second categories, `max` right of the box.
        expect(x(textNamed(bar3D({ catAxisCrossesAt: 2 } as PptxGenJS.IChartOpts), "0"))).toBeGreaterThan(x(textNamed(plain, "Q1")));
        expect(x(textNamed(bar3D({ catAxisCrossesAt: "max" } as unknown as PptxGenJS.IChartOpts), "0"))).toBeGreaterThan(x(textNamed(plain, "Q4")));
        // valAxisCrossesAt moves where the unstacked bars start, so the drawn faces change.
        expect(fronts(bar3D({ valAxisCrossesAt: 4 } as PptxGenJS.IChartOpts)).map(face => face.getAttribute("points"))).not.toEqual(fronts(plain).map(face => face.getAttribute("points")));
        // `high` puts the category labels above the value labels' top.
        const high = bar3D({ catAxisLabelPos: "high" });

        expect(y(textNamed(high, "Q1"))).toBeLessThan(y(textNamed(high, "6")));
        // The data table: its header of category names, each series' row, no axis labels, with keys and rules.
        const table = bar3D({ showDataTable: true, showDataTableKeys: true, showDataTableHorzBorder: true });

        expect(Array.from(table?.querySelectorAll(".bar3d-table") ?? [], text => text.textContent)).toEqual(["Q1", "Q2", "Q3", "Q4", "Lorem", "3", "5", "4", "7", "Ipsum", "5", "2", "6", "4"]);
        expect(textsOf(table).filter(text => text === "Q1")).toHaveLength(1);
        expect(table?.querySelectorAll(".bar3d-table-rule")).toHaveLength(2);
    });

    test("normalizes the data table, with borders and keys where the target draws them (#567)", () => {
        const options: PptxGenJS.IChartOpts = { showDataTable: true, showDataTableHorzBorder: true, showDataTableVertBorder: true, showDataTableOutline: true, showDataTableKeys: true, dataTableFontSize: 9 };

        expect(normalizeChart("bar", pair, options, "powerpoint").dataTable).toMatchObject({ horizontalBorder: true, verticalBorder: true, outline: true, keys: true, fontSize: 9 });
        // LibreOffice 26.2 draws the table's text without borders or keys (drawsDataTableBorders).
        expect(normalizeChart("bar", pair, options, "libreoffice").dataTable).toMatchObject({ horizontalBorder: false, verticalBorder: false, outline: false, keys: false });
        expect(normalizeChart("bar", pair).dataTable).toBeUndefined();
        // LibreOffice 26.2.6 draws no table on a pie, doughnut or radar chart, and a horizontal bar chart's under its
        // value labels (#583); a scatter, bubble or bar3D chart's is not drawn yet. PowerPoint draws none on a scatter or
        // bubble chart (drawsScatterDataTable off, #592) and one under a bar3D chart.
        expect(normalizeChart("pie", [pair[0]!], { showDataTable: true }).dataTable).toBeUndefined();
        expect(normalizeChart("radar", pair, { showDataTable: true }).dataTable).toBeUndefined();
        expect(normalizeChart("bar", pair, { showDataTable: true, barDir: "bar" }).dataTable).toBeDefined();
        expect(() => normalizeChart("scatter", pair, { showDataTable: true })).toThrow("showDataTable[scatter]");
        expect(() => normalizeChart("bubble", bubbleData, { showDataTable: true })).toThrow("showDataTable[bubble]");
        expect(normalizeChart("scatter", pair, { showDataTable: true }, "powerpoint").dataTable).toBeUndefined();
        expect(normalizeChart("bubble", bubbleData, { showDataTable: true }, "powerpoint").dataTable).toBeUndefined();
        // Both draw a bar3D chart's table with its borders and keys (#592).
        expect(normalizeChart("bar3D", pair, options).dataTable).toMatchObject({ horizontalBorder: true, verticalBorder: true, outline: true, keys: true });
        expect(normalizeChart("bar3D", pair, options, "powerpoint").dataTable).toMatchObject({ horizontalBorder: true, keys: true });
        // dataTableFormatCode formats the values where the target reads it (readsDataTableFormatCode), no target yet.
        expect(normalizeChart("bar", pair, { showDataTable: true, dataTableFormatCode: "0.00" }, "libreoffice").dataTable?.valueFormatCode).toBeUndefined();
        expect(normalizeChart("bar", pair, { showDataTable: true, dataTableFormatCode: "0.00" }, "powerpoint").dataTable?.valueFormatCode).toBeUndefined();
    });

    test("keeps a horizontal bar chart's category labels beside its data table (#583)", () => {
        const texts = textsOf(chartSvg("bar", pair, { showDataTable: true, barDir: "bar" }));

        // Each category on the axis and in the table's header row.
        expect(texts.filter(text => text === "Q1")).toHaveLength(2);
    });

    test("draws the radar, bar3D, scatter and secondary-axis placement options LibreOffice draws (#582)", () => {
        // A radar chart's spokes run counterclockwise for maxMin, and the value labels go on the spoke catAxisCrossesAt names.
        expect(normalizeChart("radar", quarters, { catAxisOrientation: "maxMin", catAxisCrossesAt: 2 } as unknown as PptxGenJS.IChartOpts).categoryAxis).toMatchObject({ reversed: true, valueAxisCrossing: { kind: "category", index: 1 } });
        expect(textsOf(chartSvg("radar", quarters, { catAxisLabelPos: "none" }))).not.toContain("Q1");
        expect(textsOf(chartSvg("radar", quarters, { catAxisHidden: true }))).not.toContain("Q1");
        // bar3D reverses and formats its categories.
        const bar3D = textsOf(chartSvg("bar3D", [{ name: "Lorem", labels: ["1", "2", "3"], values: [3, 5, 4] }], { catAxisOrientation: "maxMin", catLabelFormatCode: "0.0" } as unknown as PptxGenJS.IChartOpts));

        expect(bar3D.filter(text => text.endsWith(".0"))).toEqual(["1.0", "2.0", "3.0"]);
        // A scatter chart's x axis keeps its General numbers and every label, as LibreOffice draws it.
        expect(normalizeChart("scatter", pair, { catLabelFormatCode: "0.0", catAxisLabelFrequency: 2 } as unknown as PptxGenJS.IChartOpts, "powerpoint").categoryAxis).not.toHaveProperty("labelFormatCode");
        // valAxisCrossesAt and catAxisOrientation with a secondary axis; catAxisCrossesAt moves both value axes.
        const secondary = (options: PptxGenJS.IChartOpts) =>
            normalizeChart(
                [
                    { type: "bar", data: pair, options: {} },
                    { type: "line", data: [{ name: "Dolor", labels: ["Q1", "Q2", "Q3", "Q4"], values: [100, 300, 200, 400] }], options: { secondaryValAxis: true } },
                ],
                [],
                { ...options, valAxes: [{ showValAxisTitle: false }, { showValAxisTitle: false }] },
            );

        expect(secondary({ valAxisCrossesAt: 2 } as PptxGenJS.IChartOpts).valueAxes[0]?.categoryAxisCrossing).toEqual({ kind: "value", value: 2 });
        expect(secondary({ catAxisOrientation: "maxMin" } as unknown as PptxGenJS.IChartOpts).categoryAxis.reversed).toBe(true);
        // PptxGenJS writes catAxisCrossesAt on both value axes (makeValAxis): LibreOffice 26.2.6 draws both at that
        // category, overlapping (#591).
        expect(secondary({ catAxisCrossesAt: 2 } as PptxGenJS.IChartOpts).categoryAxis.valueAxisCrossing).toEqual({ kind: "category", index: 1 });
    });

    test("lays a date axis out as a time line, as LibreOffice 26.2.6 does (#580)", () => {
        const dates = (labels: string[]) => [{ name: "Lorem", labels, values: labels.map((_, index) => index + 1) }];
        const months = ["45292", "45323", "45352", "45383", "45413", "45444"];

        // Monthly serials take a month per slot; a date format formats each.
        expect<unknown>(normalizeChart("line", dates(months), { catLabelFormatCode: "mmm-yy" }).categoryAxis.dateSlots).toEqual({ labels: ["Jan-24", "Feb-24", "Mar-24", "Apr-24", "May-24", "Jun-24"], categorySlots: [0, 1, 2, 3, 4, 5] });
        // Daily serials with gaps leave the days between them empty.
        expect<unknown>(normalizeChart("bar", dates(["45292", "45295", "45297"]), { catLabelFormatCode: "d-mmm" }).categoryAxis.dateSlots).toEqual({
            labels: ["1-Jan", "2-Jan", "3-Jan", "4-Jan", "5-Jan", "6-Jan"],
            categorySlots: [0, 3, 5],
        });
        // catAxisMajorUnit labels every nth unit from the first; catAxisMinVal and catAxisMaxVal trim the time line.
        expect(normalizeChart("line", dates(months), { catLabelFormatCode: "mmm-yy", catAxisMajorUnit: 2, catAxisMajorTimeUnit: "months" }).categoryAxis.dateSlots?.labels).toEqual(["Jan-24", "", "Mar-24", "", "May-24", ""]);
        expect(normalizeChart("line", dates(months), { catLabelFormatCode: "mmm-yy", catAxisMajorUnit: 1, catAxisMajorTimeUnit: "years" }).categoryAxis.dateSlots?.labels).toEqual(["Jan-24", "", "", "", "", ""]);
        expect<unknown>(normalizeChart("line", dates(months), { catLabelFormatCode: "mmm-yy", catAxisMinVal: 45323, catAxisMaxVal: 45413 }).categoryAxis.dateSlots).toEqual({
            labels: ["Feb-24", "Mar-24", "Apr-24", "May-24"],
            categorySlots: [null, 0, 1, 2, 3, null],
        });
        // catAxisBaseTimeUnit days puts monthly serials a day per slot; years 2-digit, full names and weekdays format.
        expect(normalizeChart("line", dates(months.slice(0, 2)), { catLabelFormatCode: "dddd d mmmm yyyy", catAxisBaseTimeUnit: "days" }).categoryAxis.dateSlots?.labels).toHaveLength(32);
        expect(normalizeChart("line", dates(["45292"]), { catLabelFormatCode: 'ddd dd/mm/yy "Q"' }).categoryAxis.dateSlots?.labels).toEqual(["Mon 01/01/24 Q"]);
        // Consecutive serials with a number format stay the categories, formatted as numbers.
        expect(normalizeChart("line", dates(["1", "2", "3"]), { catLabelFormatCode: "0.0" }).categoryAxis).toMatchObject({ labelFormatCode: "0.0" });
        expect(normalizeChart("line", dates(["1", "2", "3"]), { catLabelFormatCode: "0.0" }).categoryAxis.dateSlots).toBeUndefined();
        // Several dates in one slot, a bad time unit, a fractional major unit.
        expect(() => normalizeChart("line", dates(["45292", "45293"]), { catLabelFormatCode: "mmm-yy", catAxisBaseTimeUnit: "months" })).toThrow("catAxisBaseTimeUnit[several dates in one month]");
        expect(() => normalizeChart("line", dates(months), { catLabelFormatCode: "mmm-yy", catAxisBaseTimeUnit: "weeks" })).toThrow("catAxisBaseTimeUnit");
        expect(() => normalizeChart("line", dates(months), { catLabelFormatCode: "mmm-yy", catAxisMajorUnit: 1.5 })).toThrow("whole number");
        expect(() => normalizeChart("bar3D", dates(months), { catAxisMinVal: 1 })).toThrow("catAxisMinVal[bar3D]");
        // A data table on a slot per category names its columns by date (#593); one with empty slots is not measured.
        expect(normalizeChart("line", dates(months), { catLabelFormatCode: "mmm-yy", showDataTable: true }).dataTable).toBeDefined();
        expect(() => normalizeChart("line", dates(["45292", "45352"]), { catLabelFormatCode: "mmm-yy", showDataTable: true })).toThrow("showDataTable[date axis with empty slots]");
        // On a text axis catAxisMinVal, catAxisMaxVal and catAxisMultiLevelLabels change nothing in LibreOffice.
        expect(normalizeChart("bar", quarters, { catAxisMinVal: 2, catAxisMaxVal: 3, catAxisMultiLevelLabels: true }).categoryAxis.dateSlots).toBeUndefined();
        // A scatter chart's x axis takes the bounds and the step.
        expect(normalizeChart("scatter", pair, { catAxisMinVal: 0, catAxisMaxVal: 8, catAxisMajorUnit: 2 }).categoryAxis.xScale).toEqual({ minimum: 0, maximum: 8, majorUnit: 2 });

        // The drawn axis has a category per slot, the empty ones unlabelled, and each point in its slot.
        const texts = textsOf(chartSvg("line", dates(["45292", "45295"]), { catLabelFormatCode: "d-mmm" }));

        for (const label of ["1-Jan", "2-Jan", "3-Jan", "4-Jan"]) {
            expect(texts).toContain(label);
        }
    });

    test("draws the bar3D series gridlines and accepts the depth-axis options LibreOffice does not draw (#581)", () => {
        const options = {
            serGridLine: { color: "FF0000", size: 2 },
            serAxisLabelFrequency: "2",
            serAxisLineShow: true,
            serAxisLineColor: "00FF00",
            serAxisLabelPos: "high",
            serLabelFormatCode: "0.0",
            serAxisMajorUnit: 2,
            serAxisBaseTimeUnit: "days",
        } as PptxGenJS.IChartOpts;
        const chart = normalizeChart("bar3D", pair, options);

        expect(chart.seriesAxis).toMatchObject({ gridLineColor: "#FF0000", gridLineWidth: 2, gridLineStyle: "solid" });
        // LibreOffice 26.2.6 names every series whatever serAxisLabelFrequency says (readsSeriesLabelSkip).
        expect(chart.seriesAxis?.labelSkip).toBeUndefined();
        expect(normalizeChart("bar3D", pair).seriesAxis?.gridLineStyle).toBe("none");
        expect(chartSvg("bar3D", pair, options)?.querySelectorAll(".bar3d-series-grid").length).toBe(3);
        expect(() => normalizeChart("bar3D", pair, { serAxisLabelFrequency: "0" } as PptxGenJS.IChartOpts)).toThrow("serAxisLabelFrequency");
        expect(() => normalizeChart("bar3D", pair, { serGridLine: { color: "FF0000", lorem: 1 } } as unknown as PptxGenJS.IChartOpts)).toThrow("serGridLine.lorem");
    });

    test("draws the data table's rows in place of the category labels", () => {
        const texts = textsOf(chartSvg("bar", pair, { showDataTable: true }));

        for (const text of ["Lorem", "Ipsum", "Q1", "Q4", "3", "7", "6"]) {
            expect(texts).toContain(text);
        }

        // Each category once, in the table's header row.
        expect(texts.filter(text => text === "Q1")).toHaveLength(1);
    });

    test("normalizes and draws inverted colors, overlap, the data border and label backgrounds", () => {
        expect<string | undefined>(normalizeChart("bar", quarters, { invertedColors: ["FF0000"] }).series[0]?.pointFills?.[1]).toBe("#FF0000");
        expect(normalizeChart("bar", pair, { barOverlapPct: 50 }).barOverlapPercent).toBe(50);
        expect(normalizeChart("bar", pair, { barOverlapPct: 500 }).barOverlapPercent).toBe(100);
        expect(normalizeChart("bar", pair, { barOverlapPct: 50, barGrouping: "stacked" }).barOverlapPercent).toBeUndefined();
        expect<unknown>(normalizeChart("bar", pair, { dataBorder: { pt: 3, color: "FF0000" } }).dataBorder).toEqual({ color: "#FF0000", width: 3 });
        // PptxGenJS writes a missing border color as F9F9F9.
        expect<unknown>(normalizeChart("pie", [pair[0]!], { dataBorder: { pt: 2 } } as PptxGenJS.IChartOpts).dataBorder).toEqual({ color: "#F9F9F9", width: 2 });
        expect(normalizeChart("bar", pair, { showValue: true, dataLabelBkgrdColors: true }).dataLabels.seriesColorBackground).toBe(true);

        const inverted = chartSvg("bar", quarters, { invertedColors: ["FF0000"], dataBorder: { pt: 3, color: "00B050" } });

        expect(inverted?.querySelector('[fill="#FF0000"]')).not.toBeNull();
        // Each point's `<c:dPt>` replaces the series' shape properties, line included.
        expect(inverted?.querySelector('[stroke="#00B050"]:not([stroke-width="0"])')).toBeNull();
        expect(chartSvg("bar", pair, { dataBorder: { pt: 3, color: "00B050" } })?.querySelector('[stroke="#00B050"]')).not.toBeNull();
    });

    test("places and turns the title", () => {
        const placed = normalizeChart("bar", quarters, { x: 1, y: 1, showTitle: true, title: "Lorem", titlePos: { x: 1, y: 1 }, titleRotate: 30 });

        expect(placed.titleRotate).toBe(30);
        expect(placed.titlePosition?.x).toBeGreaterThan(0);
        expect(placed.titlePosition?.y).toBeGreaterThan(0);

        const title = Array.from(chartSvg("bar", quarters, { showTitle: true, title: "Lorem title", titleRotate: 30 })?.querySelectorAll("text") ?? []).find(text => text.textContent === "Lorem title");

        expect(title?.getAttribute("transform")).toContain("matrix(0.866,0.5,");
    });

    test("draws a bar3D chart's depth-axis options", () => {
        const svg = chartSvg("bar3D", pair, { barGrouping: "standard", showSerAxisTitle: true, serAxisTitle: "Series", serAxisLabelColor: "FF0000" });

        expect(svg?.querySelector(".bar3d-series-title")?.textContent).toBe("Series");
        expect(svg?.querySelector(".bar3d-series-label")?.getAttribute("fill")).toBe("#FF0000");
        expect(chartSvg("bar3D", pair, { barGrouping: "standard", serAxisHidden: true })?.querySelector(".bar3d-series-label")).toBeNull();
        expect(normalizeChart("bar3D", pair, { serAxisOrientation: "maxMin" } as PptxGenJS.IChartOpts).seriesAxis?.reversed).toBe(true);
        expect(normalizeChart("bar", pair).seriesAxis).toBeUndefined();
    });

    test("scales a bubble chart's x axis to its x values, not its bubbles (#548)", () => {
        const texts = textsOf(
            chartSvg(
                "bubble",
                [
                    { name: "X", values: [1, 2, 3, 4] },
                    { name: "Lorem", values: [3, 5, 4, 7], sizes: [1, 5, 10, 20] },
                ],
                {},
            ),
        );

        // LibreOffice 26.2 numbers x values 1..4 from 0.5 to 4.5 by 0.5, whatever the sizes.
        expect(texts).toContain("0.5");
        expect(texts).toContain("4.5");
    });

    test("draws what PowerPoint draws for the charts measured in #565 and #604", () => {
        const pieData = [{ name: "Share", labels: ["Lorem", "Ipsum"], values: [40, 60] }];
        const plotArea = { fill: { color: "FFF2CC" }, border: { color: "0000FF", pt: 3 } };
        const page = (target: "libreoffice" | "powerpoint", add: (slide: PptxGenJS.Slide) => void) => {
            const presentation = new PuppeteerGen(undefined, { target });

            add(presentation.addSlide() as unknown as PptxGenJS.Slide);

            return { model: presentation.getModel(), svgs: Array.from(renderedPage(presentation).querySelectorAll(".slide-chart svg")) };
        };

        // A pie's series name on its own line above the value, on both targets.
        for (const target of ["libreoffice", "powerpoint"] as const) {
            const { svgs } = page(target, slide => slide.addChart("pie", pieData, { x: 0, y: 0, w: 4, h: 3, showSerName: true, showValue: true }));

            expect(svgs[0]?.textContent).toContain("Share");
        }

        // The plot area's fill and border on a pie, doughnut, radar and bar3D chart: drawn on PowerPoint, rejected on
        // LibreOffice.
        const drawn = page("powerpoint", slide => {
            slide.addChart("pie", pieData, { x: 0, y: 0, w: 3, h: 3, plotArea });
            slide.addChart("doughnut", pieData, { x: 3, y: 0, w: 3, h: 3, fill: "FFF2CC", border: { color: "0000FF", pt: 3 } });
            slide.addChart("radar", categoricalData, { x: 6, y: 0, w: 3, h: 3, plotArea });
            slide.addChart("bar3D", categoricalData, { x: 0, y: 3, w: 3, h: 3, plotArea });
        });

        for (const svg of drawn.svgs) {
            expect(svg.querySelector('[fill="#FFF2CC"][stroke="#0000FF"]')).not.toBeNull();
        }

        expect(() => normalizeChart("doughnut", pieData, { plotArea }, "libreoffice")).toThrow("plotArea[doughnut]");

        // A bar3D value axis line and its tick marks: each step's major tick and each fifth's minor one on PowerPoint.
        const axis = page("powerpoint", slide => slide.addChart("bar3D", categoricalData, { x: 0, y: 0, w: 5, h: 4, valAxisLineColor: "C00000", valAxisMajorTickMark: "cross", valAxisMinorTickMark: "outside", valAxisMaxVal: 6, valAxisMajorUnit: 2 }));
        const bar3D = axis.svgs[0];

        expect(bar3D?.querySelector(".bar3d-value-axis")?.getAttribute("stroke")).toBe("#C00000");
        expect(bar3D?.querySelectorAll(".bar3d-value-tick")).toHaveLength(4 + 16);
        expect(page("powerpoint", slide => slide.addChart("bar3D", categoricalData, { x: 0, y: 0, w: 5, h: 4, valAxisLineShow: false, valAxisMajorTickMark: "none" })).svgs[0]?.querySelector(".bar3d-value-axis, .bar3d-value-tick")).toBeNull();

        // A combo line type's markers take PptxGenJS's written fill, chartColors[index of the series in the chart], on
        // PowerPoint only; a line chart alone writes its line's color.
        const combo = (target: "libreoffice" | "powerpoint") => {
            const chart = normalizeChart(
                [
                    { type: "bar", data: categoricalData, options: {} },
                    { type: "line", data: [{ name: "Dolor", labels: ["A", "B", "C"], values: [6, 9, 7] }], options: {} },
                ],
                [],
                { chartColors: ["C0504D", "4F81BD", "9BBB59"] },
                target,
            );

            return chart.series.map(series => series.markerFill);
        };

        expect<(string | undefined)[]>(combo("powerpoint")).toEqual([undefined, undefined, "#9BBB59"]);
        expect(combo("libreoffice")).toEqual([undefined, undefined, undefined]);
        expect(normalizeChart("line", categoricalData, {}, "powerpoint").series.map(series => series.markerFill)).toEqual([undefined, undefined]);
    });
});
