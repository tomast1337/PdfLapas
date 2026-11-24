import { describe, expect, test } from "bun:test";
import PptxGenJSRuntime from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";

// A function: PptxGenJS writes onto the data too (labels become arrays of levels).
function data(): PptxGenJS.OptsChartData[] {
    return [{ name: "Lorem", labels: ["Q1", "Q2", "Q3", "Q4"], values: [3, 5, 4, 7] }];
}

function lineOptions(): PptxGenJS.IChartOpts {
    return { x: 1, y: 1, w: 6, h: 4, showLegend: false, lineDataSymbolSize: 14, lineDataSymbolLineSize: 2, lineDataSymbolLineColor: "000000" };
}

/** The deck archive PptxGenJS 4.0.1 writes, stored, not deflated, so its XML reads from the bytes without a zip reader. */
async function pptxGenJsArchive(add: (pptx: PptxGenJSRuntime) => void): Promise<string> {
    const pptx = new PptxGenJSRuntime();

    add(pptx);

    return Buffer.from((await pptx.write({ outputType: "nodebuffer", compression: false })) as Uint8Array).toString("utf8");
}

describe("a chart options object shared between charts (#638)", () => {
    test("PptxGenJS 4.0.1 writes every chart on one object from its final state", async () => {
        const markerWidths = [];

        for (const calls of [1, 2, 3]) {
            const options = lineOptions();
            const archive = await pptxGenJsArchive(pptx => {
                for (let call = 0; call < calls; call++) {
                    pptx.addSlide().addChart("line", data(), options);
                }
            });
            const widths = [];

            for (const [, width] of archive.matchAll(/<c:marker>[\s\S]*?<a:ln w="(\d+)"/g)) {
                widths.push(Number(width));
            }

            markerWidths.push(widths);
        }

        // 2pt is 25400 EMU; each call multiplies the stored width by 12700 again, past ST_LineWidth's 20116800 on the
        // second, and every chart writes the last value.
        expect(markerWidths).toEqual([[25400], [322580000, 322580000], [4096766000000, 4096766000000, 4096766000000]]);

        const mixed = await pptxGenJsArchive(pptx => {
            const options: PptxGenJS.IChartOpts = { x: 1, y: 1, w: 6, h: 4, showPercent: true, showValue: true };

            pptx.addSlide().addChart("pie", data(), options);
            pptx.addSlide().addChart("bar", data(), options);
        });

        // The pie is written as a bar chart, as the second call left the object.
        expect(mixed.match(/<c:barChart>/g)?.length).toBe(2);
        expect(mixed.includes("<c:pieChart>")).toBe(false);
    });

    test("draws each chart from its options at its own call and warns once, without writing onto the object", async () => {
        const reused = new PuppeteerGen();
        const copied = new PuppeteerGen();
        const options = lineOptions();

        for (const type of ["line", "line", "bar"] as const) {
            reused.addSlide().addChart(type, data(), options);
            copied.addSlide().addChart(type, data(), { ...lineOptions() });
        }

        expect(reused.getModel()).toEqual(copied.getModel());
        expect(options).toEqual(lineOptions());
        expect(copied.warnings()).toEqual([]);
        expect(reused.warnings()).toHaveLength(1);
        expect(reused.warnings()[0]).toContain("addChart(): the options object was given to an earlier addChart() too");
        expect(reused.warnings()[0]).toContain("{ ...options }");
    });
});

describe("options objects PptxGenJS writes stable values onto (#638)", () => {
    /** A text box, a table and a shape on two slides each: on one options object per kind, or on fresh copies. */
    function addTwice(pptx: { addSlide: () => PptxGenJS.Slide }, reuse: boolean): void {
        const text: PptxGenJS.TextPropsOptions = { x: 0.5, y: 0.5, w: 3, h: 1, fontFace: "Arial", fontSize: 14 };
        const table: PptxGenJS.TableProps = { x: 0.5, y: 2, w: 4, h: 1, fontFace: "Arial", fontSize: 12, border: { type: "solid", pt: 1, color: "000000" } };
        const shape: PptxGenJS.ShapeProps = { x: 5, y: 0.5, w: 2, h: 1, line: { color: "FF0000", width: 2 } };

        for (let index = 0; index < 2; index++) {
            const slide = pptx.addSlide();

            slide.addText("Lorem", reuse ? text : { ...text });
            slide.addTable([[{ text: "Lorem" }, { text: "Ipsum" }]], reuse ? table : { ...table });
            slide.addShape("rect", reuse ? shape : { ...shape, line: { color: "FF0000", width: 2 } });
        }
    }

    test("PptxGenJS 4.0.1 writes the same slide for a reused object as for copies, but for object names", async () => {
        const slides = [];

        for (const reuse of [true, false]) {
            const archive = await pptxGenJsArchive(pptx => addTwice(pptx, reuse));

            for (const [slide] of archive.matchAll(/<p:sld [\s\S]*?<\/p:sld>/g)) {
                slides.push(slide.replace(/ name="[^"]*"/g, ""));
            }
        }

        expect(slides).toHaveLength(4);
        expect(slides.slice(0, 2)).toEqual(slides.slice(2));
    });

    test("draws a reused object as its copies, without warning", () => {
        const models = [];

        for (const reuse of [true, false]) {
            const pres = new PuppeteerGen();

            addTwice(pres, reuse);
            expect(pres.warnings()).toEqual([]);
            models.push(pres.getModel());
        }

        // `as`: the loop pushed two models, as the length check shows.
        expect(models).toHaveLength(2);
        expect(models[0]).toEqual(models[1] as NonNullable<(typeof models)[1]>);
    });
});
