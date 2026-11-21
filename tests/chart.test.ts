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
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { unresolvedModelHtml } from "../packages/pptxgenjs/src/render/document";

const categoricalData: PptxGenJS.OptsChartData[] = [
    { name: "First", labels: ["A", "B", "C"], values: [2, 5, 3] },
    { name: "Second", labels: ["A", "B", "C"], values: [4, 1, 6] },
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
            normalizeChart("bar", categoricalData, { showSerName: true });
            throw new Error("Expected normalization to fail");
        } catch (error) {
            expect(error).toBeInstanceOf(UnsupportedChartError);
            expect((error as UnsupportedChartError).code).toBe("UNSUPPORTED_CHART");
            expect((error as UnsupportedChartError).reason).toBe("chart-option");
            expect((error as UnsupportedChartError).unsupportedOptions).toEqual(["showSerName"]);
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
        expect(normalizeChart("pie", [categoricalData[0]!]).dataLabels.position).toBe("bestFit");
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

    test("rejects data labels it cannot draw, naming the option and the reason", () => {
        const unsupported = (type: PptxGenJS.CHART_NAME, data: PptxGenJS.OptsChartData[], options: PptxGenJS.IChartOpts) => {
            try {
                normalizeChart(type, data, options);
            } catch (error) {
                return (error as UnsupportedChartError).unsupportedOptions;
            }

            throw new Error("Expected normalization to fail");
        };

        expect(unsupported("area", categoricalData, { showValue: true })).toEqual(["showValue[area]"]);
        expect(unsupported("radar", categoricalData, { showValue: true })).toEqual(["showValue[radar]"]);
        expect(unsupported("pie", [categoricalData[0]!], { showValue: true, dataLabelPosition: "outEnd" })).toEqual(["dataLabelPosition[outEnd]"]);
        expect(unsupported("bar", categoricalData, { showValue: true, dataLabelFormatCode: "0;[Red]-0" })).toEqual(["dataLabelFormatCode[format]"]);
        // Styling without a shown label draws nothing in PptxGenJS's output either.
        expect(normalizeChart("area", categoricalData, { dataLabelColor: "FF0000", showLabel: true }).dataLabels.showValue).toBe(false);
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
        expect(() => normalizeChart("bar3D", categoricalData, { catAxisLabelColor: "FF0000" })).toThrow("catAxisLabelColor[bar3D]");
        expect(() => normalizeChart("radar", categoricalData, { showCatAxisTitle: true })).toThrow("showCatAxisTitle[radar]");
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
        expect(() => normalizeChart("radar", categoricalData, { layout: { x: 0.1 } })).toThrow("layout[radar]");
        expect(() => normalizeChart("bar", categoricalData, { plotArea: { border: { color: "FF0000", dashType: "dash" } as PptxGenJS.BorderProps } })).toThrow("plotArea.border.dashType");
        // The deprecated top-level fill and border stay unsupported.
        expect(() => normalizeChart("bar", categoricalData, { fill: "FF0000" })).toThrow(UnsupportedChartError);
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
        expect(normalizeChart("bar3D", categoricalData, { barGrouping: "stacked" }).valueAxisMaximum).toBe(10);
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

        presentation.addSlide().addChart("bar3D", [{ ...categoricalData[0], name: tricky }], { x: 0, y: 0, w: 5, h: 3, showTitle: true, title: tricky, showLegend: true });
        const svg = renderedPage(presentation).querySelector(".slide-chart svg");

        expect(svg?.querySelector(".bar3d-title")?.textContent).toBe(tricky);
        expect(Array.from(svg?.querySelectorAll(".bar3d-series-label, .bar3d-legend") ?? [], label => label.textContent)).toEqual([tricky, tricky]);
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

    test("sizes the automatic plot-area margins from the text, not the chart height", () => {
        const plotBounds = (hInches: number, options: PptxGenJS.IChartOpts) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart("bar", categoricalData, { x: 0, y: 0, w: 5, h: hInches, ...options });
            const rows = Array.from(renderedPage(presentation).querySelectorAll(".slide-chart path"), path => path.getAttribute("d")?.match(/^M[\d.]+ ([\d.]+)L[\d.]+ \1$/)?.[1]).filter(y => y !== undefined);
            const ys = rows.map(Number);

            return { top: Math.min(...ys), bottom: hInches * 96 - Math.max(...ys) };
        };

        // Below: 2% gap, 10pt legend line (1.2em) and 2% gap, 5px ticks, 8px label margin, 12pt label line (1.2em).
        // Above: 2% gap and half of the 12pt value label. Within 1px: zrender snaps lines to half pixels.
        for (const h of [3, 4.5]) {
            const gap = h * 96 * 0.02;
            const withLegend = plotBounds(h, { showLegend: true, legendPos: "b" });

            expect(Math.abs(withLegend.bottom - (2 * gap + 16 + 5 + 8 + 19.2))).toBeLessThanOrEqual(1);
            expect(Math.abs(withLegend.top - (gap + 8))).toBeLessThanOrEqual(1);
            expect(Math.abs(plotBounds(h, {}).bottom - (gap + 5 + 8 + 19.2))).toBeLessThanOrEqual(1);
        }
    });

    test("draws a point marker as PptxGenJS's c:marker writes it: size in points, 0.75pt outline, none on area series", () => {
        const markers = (type: "area" | "line" | "radar" | "scatter", options: PptxGenJS.IChartOpts = {}) => {
            const presentation = new PuppeteerGen();

            presentation.addSlide().addChart(type, categoricalData, { x: 0, y: 0, w: 5, h: 3, ...options });

            return Array.from(renderedPage(presentation).querySelectorAll('.slide-chart path[ecmeta_ssr_type="chart"][ecmeta_data_index]'));
        };
        const lineMarkers = markers("line");
        // The default 6pt circle: a unit circle scaled to a 4px radius (8px across).
        const scale = pointsToPixels(PPTX_DEFAULTS.chart.lineDataSymbolSize) / 2;

        expect(lineMarkers).toHaveLength(6);
        expect(lineMarkers[0]?.getAttribute("transform")).toStartWith(`matrix(${scale},0,0,${scale},`);
        expect(lineMarkers[0]?.getAttribute("stroke")).toBe(lineMarkers[0]?.getAttribute("fill") ?? "fill");
        // ECharts writes the stroke inside the scaled transform.
        expect(Number(lineMarkers[0]?.getAttribute("stroke-width")) * scale).toBe(pointsToPixels(PPTX_DEFAULTS.chart.lineDataSymbolLineSizePt));
        expect(markers("line", { lineDataSymbol: "square" })).toHaveLength(6);
        expect(markers("line", { lineDataSymbol: "none" })).toHaveLength(0);
        expect(markers("area")).toHaveLength(0);
        expect(markers("radar", { radarStyle: "marker" })).toHaveLength(6);
        expect(markers("radar")).toHaveLength(0);
        expect(markers("scatter", { lineDataSymbol: "none" })).toHaveLength(0);
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

    test("draws outside major tick marks on the category axis, as PptxGenJS writes them", () => {
        // PptxGenJS 4.0.1: `<c:majorTickMark val="${catAxisMajorTickMark || 'out'}"/>` on every category axis but scatter's.
        expect(normalizeChart("bar", categoricalData).categoryMajorTickMark).toBe("outside");
        const presentation = new PuppeteerGen();

        presentation.addSlide().addChart("bar", categoricalData, { x: 0, y: 0, w: 5, h: 3, valAxisMajorTickMark: "none" });
        const paths = Array.from(renderedPage(presentation).querySelectorAll(".slide-chart path"), path => path.getAttribute("d") ?? "");
        // A vertical 5px stroke at each of the four category boundaries of three categories.
        const ticks = paths.filter(d => {
            const match = d.match(/^M([\d.]+) ([\d.]+)L([\d.]+) ([\d.]+)$/);

            return match !== null && match[1] === match[3] && Number(match[4]) - Number(match[2]) === 5;
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

    test("validates logarithmic axes and explicitly rejects unsupported format and cross ticks", () => {
        const logarithmic = normalizeChart("line", [{ name: "Growth", labels: ["A", "B"], values: [1, 100] }], {
            valAxisMinVal: 1,
            valAxisMaxVal: 100,
            valAxisLogScaleBase: 10,
        });

        expect(logarithmic.valueAxes[0]?.logScaleBase).toBe(10);
        expect(() => normalizeChart("line", categoricalData, { valAxisLogScaleBase: 10 })).toThrow("positive minimum");
        expect(() => normalizeChart("line", categoricalData, { valAxisLogScaleBase: 1 })).toThrow("between 2 and 99");
        expect(() => normalizeChart("line", categoricalData, { valAxisLabelFormatCode: "0;[Red]-0" })).toThrow(UnsupportedChartError);
        expect(() => normalizeChart("line", categoricalData, { valAxisMajorTickMark: "cross" })).toThrow(UnsupportedChartError);
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
            expect((error as UnsupportedChartError).unsupportedOptions).toEqual(["series[1].options.secondaryCatAxis"]);
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
