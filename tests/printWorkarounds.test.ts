import { afterAll, describe, expect, test } from "bun:test";
import { pdfContentStreams } from "./pdfContentStreams";
import { JSDOM } from "jsdom";
import type PptxGenJS from "pptxgenjs";
import { closeBrowser } from "../packages/puppeteer/src/browser";
import { PuppeteerGen, PuppeteerGenPowerPoint } from "../packages/pptxgenjs/src/node";
import { PRINT_ZOOM, printHtml } from "../packages/puppeteer/src/printWorkarounds";
import { PuppeteerPdfRenderer } from "../packages/puppeteer/src/puppeteerPdf";
import { inchesToPixels, inchesToPoints, pixelsToInches } from "@pdflapas/model";

/** The document Puppeteer prints, parsed for structural assertions. */
async function printedPage(presentation: PuppeteerGen): Promise<Document> {
    return new JSDOM(printHtml((await presentation.toHtml()).html)).window.document;
}

afterAll(() => closeBrowser());

describe("Puppeteer print workarounds", () => {
    test("moves each horizontal paragraph so the strut's centred baseline lands on data-line-baseline", async () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addText("Horizontal", { x: 1, y: 1, w: 4, h: 1 });
        slide.addText("Vertical", { x: 1, y: 3, w: 1, h: 3, vert: "vert270" });
        const [horizontal, vertical] = (await printedPage(presentation)).querySelectorAll<HTMLElement>(".text-paragraph");

        // 18pt (24px) default: a 28.8px line box with its baseline at 24px. The strut's sits 115 of the 230.39 zoomed
        // layout px of the line below its top (half, floored to a whole px), 14.375px, so the paragraph moves 9.625px.
        expect(horizontal?.style.position).toBe("relative");
        expect(parseFloat(horizontal?.style.top ?? "")).toBeCloseTo(9.625, 9);
        expect([vertical?.style.position, vertical?.style.top]).toEqual(["", ""]);
    });

    test("prints each baseline at data-line-baseline, within half a print device px", async () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const boxes = [
            { fontSize: 6, y: 0.3137 },
            { fontSize: 7, y: 0.9071 },
            { fontSize: 9.5, y: 1.5513 },
            { fontSize: 12, y: 2.2299 },
            { fontSize: 14, y: 2.8642 },
            { fontSize: 18, y: 3.6085 },
        ];

        for (const { fontSize, y } of boxes) {
            slide.addText("Hxg", { x: 1, y, w: 3, h: 0.6, fontSize, fontFace: "Arial", margin: 0, valign: "top" });
        }

        const baselines = [...(await presentation.toHtml()).html.matchAll(/data-line-baseline="([^"]*)"/g)].map(match => Number(match[1]));
        // The text baselines of the one-page Chromium PDF, in drawing order and PRINT_ZOOM device px from the page top
        // (each `Tm`'s y).
        const printed: number[] = [];

        for (const content of pdfContentStreams(await presentation.toPdf())) {
            for (const match of content.matchAll(/(?:^|\n)1 0 0 -1 [\d.]+ ([\d.]+) Tm(?=\n)/g)) {
                printed.push(Number(match[1]));
            }
        }

        // Printed less intended, in pt: Chromium paints a baseline at a whole device px, 0.09375pt, so half of it at most.
        const errors = boxes.map(({ y }, index) => inchesToPoints(pixelsToInches((printed[index] ?? NaN) / PRINT_ZOOM - inchesToPixels(y) - (baselines[index] ?? NaN))));

        for (const error of errors) {
            expect(Math.abs(error)).toBeLessThanOrEqual(0.05);
        }
    }, 60_000);

    test("undoes Chromium's whole-px SVG root snap, in the object's pre-transform space (issue #42)", async () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addShape("rect", { x: 1, y: 1, w: 1, h: 1 });
        // 6.55in, 1.6in -> 628.8px, 153.6px: Chromium draws the SVG from the nearest 1 / PRINT_ZOOM px, 628.75, 153.625.
        slide.addShape("line", { x: 6.55, y: 1.6, w: 2.6, h: 0, line: { color: "4472C4", width: 2.25 } });
        slide.addText("Shape text", { x: 6.55, y: 1.6, w: 2, h: 1, shape: "rect" });
        // Chromium applies the page-space difference after the object's rotation and flips.
        slide.addShape("rect", { x: 6.55, y: 1.6, w: 1, h: 1, rotate: 90 });
        slide.addShape("rect", { x: 6.55, y: 1.6, w: 1, h: 1, flipH: true });
        slide.addShape("rect", { x: 6.55, y: 1.6, w: 1, h: 1, rotate: 180, flipH: true, flipV: true });
        const offsets = Array.from((await printedPage(presentation)).querySelectorAll<SVGSVGElement>(".shape-svg"), svg => svg.style.transform);

        expect(offsets).toEqual(["", "translate(0.05px, -0.025px)", "translate(0.05px, -0.025px)", "translate(-0.025px, -0.05px)", "translate(-0.05px, -0.025px)", "translate(0.05px, -0.025px)"]);
    });

    // PptxGenJS writes <a:ext cx="0"> and PowerPoint and LibreOffice draw the stroke; Chromium paints no SVG root 0px wide.
    test("prints a zero-width line as a stroke at its x, from its y over its height", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addShape("line", { x: 3.13, y: 1.07, w: 0, h: 2, line: { color: "2CA8E2", width: 1 } });
        const content = pdfContentStreams(await presentation.toPdf()).join("\n");
        // The SVG's `scale 0 0 scale x y cm`, then (no nested q) its stroked path in CSS px.
        const [, scale, x, y, length] = content.match(/\n([\d.]+) 0 0 \1 ([\d.]+) ([\d.]+) cm\n[^q]*?\n0 0 m\n0 ([\d.]+) l\nS\n/) ?? [];

        expect([Number(x) / Number(scale), Number(y) / Number(scale), Number(length)].map(value => value.toFixed(3))).toEqual([inchesToPixels(3.13), inchesToPixels(1.07), inchesToPixels(2)].map(value => value.toFixed(3)));
    }, 60_000);

    // A cell has no autofit (PptxGenJS writes a bare <a:bodyPr/>), so hand-written cell shrink fails, as in WASM (#197).
    test("fails on a table cell with a text fit other than none", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addTable([[{ text: "Cell" }]], { x: 1, y: 1, w: 2, objectName: "Grid" });
        const html = (await presentation.toHtml()).html;
        const shrink = html.replace('class="table-cell-text" data-text-fit="none"', 'class="table-cell-text" data-text-fit="shrink" data-text-fit-step="0.95" data-text-fit-max-steps="20" data-text-fit-tolerance="0.5"');

        expect(shrink).not.toBe(html);
        expect(() => printHtml(html)).not.toThrow();
        expect(() => printHtml(shrink)).toThrow(`Cannot print 'Grid': its .table-cell-text needs data-text-fit="none"`);
    });

    test("draws a dash pattern that starts with a gap as the same pattern shifted by an offset", async () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addShape("line", { x: 1, y: 1, w: 4, h: 0, line: { color: "ED7D31", width: 2.25, dashType: "dash", beginArrowType: "oval" } });
        slide.addShape("line", { x: 1, y: 2, w: 4, h: 0, line: { color: "ED7D31", width: 2, dashType: "dash" } });
        const [gapFirst, plain] = (await printedPage(presentation)).querySelectorAll(".shape-geometry");
        // The model's "0 3.9 12 9 ..." loses its zero-length dash; the offset leaves the first 3.9px blank.
        const parts = gapFirst?.getAttribute("stroke-dasharray")?.split(" ").map(Number) ?? [];

        expect(parts.slice(0, 2)).toEqual([12, expect.closeTo(9, 9)]);
        let period = 0;

        for (const part of parts) {
            period += part;
        }

        expect(Number(gapFirst?.getAttribute("stroke-dashoffset"))).toBeCloseTo(period - 3.9, 9);
        expect(plain?.getAttribute("stroke-dasharray")).toBe(presentation.getModel().slides[0]?.objects[1]?.style.line?.dashArray ?? "");
        expect(plain?.hasAttribute("stroke-dashoffset")).toBe(false);
    });

    test("reads a dash pattern separated by commas and/or whitespace, as SVG allows", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addShape("line", { x: 1, y: 1, w: 4, h: 0, line: { color: "ED7D31", width: 2.25, dashType: "dash", beginArrowType: "oval" } });
        const { html } = await presentation.toHtml();
        const dashes = html.match(/stroke-dasharray="([^"]+)"/)?.[1] ?? "";
        const printed = printHtml(html);

        expect(dashes).toStartWith("0 ");
        expect(printed).toContain("stroke-dashoffset");

        for (const separator of [",", ", ", " ,", "  ", " \t", ",\n"]) {
            const written = dashes.split(" ").join(separator);

            expect(printHtml(html.replace(`stroke-dasharray="${dashes}"`, `stroke-dasharray=" ${written} "`))).toBe(printed);
        }
    });

    test("draws a text box's line as an outline centred on its edge, which leaves the text where its inset puts it", async () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addText("outlined", { x: 1, y: 1, w: 2, h: 1, line: { color: "FF0000", width: 3, dashType: "dash" } });
        slide.addText("plain", { x: 1, y: 3, w: 2, h: 1 });
        const [outlined, plain] = (await printedPage(presentation)).querySelectorAll<HTMLElement>(".slide-text");

        // 3pt = 4px: centred on the edge, from 2px outside the box to 2px inside it.
        expect(outlined?.getAttribute("style")).toEndWith("border-style: none; outline: 4px dashed #FF0000; outline-offset: -2px;");
        expect(plain?.getAttribute("style")).not.toContain("outline");
    });

    test("offsets a table by half its first row's widest top border and its first column's widest left border, which DrawingML centres on the grid", async () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const cell = (pt: number) => ({ text: "x", options: { border: { type: "solid" as const, pt, color: "000000" } } });

        slide.addTable([[cell(1), cell(3)]], { x: 1, y: 1, w: 4 });
        slide.addTable([[{ text: "x", options: { border: { type: "none" as const } } }]], { x: 1, y: 3, w: 4 });
        const [bordered, borderless] = (await printedPage(presentation)).querySelectorAll<HTMLElement>("table.slide-table");

        // The first column's left border is 1pt (1.25px snapped), the first row's widest top 3pt = 4px: Chromium puts their
        // outer halves inside the table box, so the grid moves left by 0.625px and up by 2px.
        expect([bordered?.style.position, bordered?.style.left, bordered?.style.top]).toEqual(["relative", "-0.625px", "-2px"]);
        expect([borderless?.style.position, borderless?.style.left]).toEqual(["", ""]);
    });

    test("grows each bordered cell's text box over the inner half of its borders, where DrawingML measures the margins from", async () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const table = (text: string, position: { y: number; h?: number }, textDirection?: "vert") =>
            slide.addTable([[{ text, options: { textDirection, border: [{ type: "solid", pt: 3, color: "000000" }, { type: "none" }, { type: "solid", pt: 1.5, color: "000000" }, { type: "none" }] } }]], { x: 1, w: 4, ...position });

        // Text in a row keeps its box in flow; vertical text in a row with a height fills the cell absolutely.
        table("flow", { y: 1 });
        table("absolute", { y: 3, h: 1 }, "vert");
        slide.addTable([[{ text: "none", options: { border: { type: "none" as const } } }]], { x: 1, y: 5, w: 4 });
        const [flow, absolute, none] = Array.from((await printedPage(presentation)).querySelectorAll(".table-cell-text"), text => text.getAttribute("style") ?? "");

        // 3pt = 4px on top, 1.5pt = 2px at the bottom, hidden sides 0px.
        expect(flow).toEndWith("; margin: -2px 0px -1px 0px; width: calc(100% + 0px);");
        expect(absolute).toEndWith("; inset: -2px 0px -1px 0px;");
        expect(none).not.toMatch(/margin|inset/);
    });

    test("grows a cell's text box over the half of a neighbour's wider border that Chromium collapses into it", async () => {
        const presentation = new PuppeteerGen();
        const line = { type: "solid", pt: 1.5, color: "000000" } as const;
        const none = { type: "none" } as const;

        // A bottom border only on the first row, a right border only on the first column: the cells below and to the
        // right write none, yet Chromium gives each the inner half of the shared 2px line.
        presentation.addSlide().addTable(
            [
                [
                    { text: "a", options: { border: [none, line, line, none] } },
                    { text: "b", options: { border: [none, none, line, none] } },
                ],
                [{ text: "c", options: { border: [none, line, none, none], colspan: 2 } }],
                [{ text: "d" }, { text: "e" }],
            ],
            { x: 1, y: 1, w: 4 },
        );
        const page = await printedPage(presentation);
        const [a, b, c] = Array.from(page.querySelectorAll(".table-cell-text"), text => text.getAttribute("style") ?? "");

        expect(a).toEndWith("; margin: 0px -1px -1px 0px; width: calc(100% + 1px);");
        expect(b).toEndWith("; margin: 0px 0px -1px -1px; width: calc(100% + 1px);");
        // The spanning cell takes the widest of the edges above it; its own right border is the table's outer edge.
        expect(c).toEndWith("; margin: -1px -1px 0px 0px; width: calc(100% + 1px);");
        // No top border on the first row and no left border on the first column: the table stays where it is.
        expect(page.querySelector<HTMLElement>("table.slide-table")?.getAttribute("style")).not.toContain("position: relative");
    });

    // PowerPoint keeps 0.2667in rows of 10pt text (0.05in margins, filling the row) 19.2pt apart whatever the borders;
    // Chromium grew each row by the half of the border above it that the cell itself does not write.
    test("prints rows at their height when the borders between them belong to one side only", async () => {
        const presentation = new PuppeteerGen();
        const line = { type: "solid", pt: 0.5, color: "000000" } as const;
        const none = { type: "none" } as const;
        const rows = (count: number): PptxGenJS.TableRow[] => Array.from({ length: count }, (_, index) => [{ text: `Row ${index + 1}`, options: { fill: { color: index % 2 ? "FF0000" : "0000FF" }, border: [none, none, line, none] } }]);
        const slide = presentation.addSlide();
        const heights = [0.2667, 0.31, 0.2667, 0.2833, 0.2667, 0.29];

        slide.addTable(rows(12), { x: 0.5, y: 0.5, w: 3, rowH: 0.2667, fontSize: 10, fontFace: "Arial" });
        slide.addTable(rows(heights.length), { x: 5, y: 0.5, w: 3, rowH: heights, fontSize: 10, fontFace: "Arial" });
        // The red and blue cell fills of the Chromium PDF, in page order, at their top-left corner in PRINT_ZOOM device px
        // (the unit Chromium writes rectangles in, after its page `cm`).
        const fills: Array<{ x: number; y: number }> = [];

        for (const content of pdfContentStreams(await presentation.toPdf())) {
            // A rectangle filled (`re` then `f`) while the fill colour (`rg`) is red or blue.
            let colored = false;

            for (const [, color, x, y] of content.matchAll(/(?:^|\n)(?:(?:.* )?([\d.]+ [\d.]+ [\d.]+) rg|([\d.]+) ([\d.]+) [\d.]+ [\d.]+ re\nf)(?=\n)/g)) {
                if (color !== undefined) {
                    colored = color === "1 0 0" || color === "0 0 1";
                } else if (colored) {
                    fills.push({ x: Number(x), y: Number(y) });
                }
            }
        }

        // Each fill's top, in PRINT_ZOOM device px from the page top, against the exact running sum of the row heights.
        const expected = (x: number, inches: number[]) => {
            const tops: { x: number; y: number }[] = [];
            let above = 0;

            for (const h of inches) {
                tops.push({ x, y: Math.round((0.5 + above) * 96 * PRINT_ZOOM) });
                above += h;
            }

            return tops;
        };

        expect(fills).toEqual([...expected(0.5 * 96 * PRINT_ZOOM, Array(12).fill(0.2667)), ...expected(5 * 96 * PRINT_ZOOM, heights)].map(({ x, y }) => ({ x, y: expect.closeTo(y, -0.5) as unknown as number })));
    }, 60_000);

    test("measures the inset and offset from the border width Chromium lays out, floored to a whole print device px", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addTable([[{ text: "x", options: { border: { type: "solid", pt: 1, color: "000000" } } }]], { x: 1, y: 1, w: 4 });
        const page = await printedPage(presentation);

        // 1pt = 1.333px is 10.667 device px at PRINT_ZOOM 8; Chromium draws 10, so 1.25px and its half 0.625px.
        expect(page.querySelector(".table-cell-text")?.getAttribute("style")).toEndWith("; margin: -0.625px -0.625px -0.625px -0.625px; width: calc(100% + 1.25px);");
        expect(page.querySelector<HTMLElement>("table.slide-table")?.style.top).toBe("-0.625px");
    });

    test("lays the document out at the print zoom, which page.pdf() scales back", () => {
        const html = `<!DOCTYPE html><html lang="en"><head></head><body></body></html>`;

        expect(printHtml(html)).toBe(`<!DOCTYPE html><html style="zoom: ${PRINT_ZOOM};" lang="en"><head></head><body></body></html>`);
        expect(printHtml(`<!DOCTYPE html><html style="color: red"><body></body></html>`)).toBe(`<!DOCTYPE html><html style="color: red; zoom: ${PRINT_ZOOM};"><body></body></html>`);
        expect(() => printHtml("<p>fragment")).toThrow("Cannot print: <html> must be written out in the document");
    });

    test("appends Chromium's fallback list to each run's and bullet's face", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addText(
            [
                { text: "a", options: { fontFace: "Calibri Light" } },
                { text: "b", options: { fontFace: 'Odd "Face"' } },
                { text: "c", options: { fontFace: "Cambria" } },
            ],
            { x: 1, y: 1, w: 4, h: 1, bullet: true },
        );
        const page = await printedPage(presentation);
        const family = (selector: string) =>
            page
                .querySelector(selector)
                ?.getAttribute("style")
                ?.match(/font-family: ([^;]+);/)?.[1];

        // Carlito and Caladea come right after the face, before its fallback faces: metric-compatible with Calibri and Cambria, as the WASM renderer maps them.
        expect(family(".text-bullet")).toBe('"Calibri Light", "Carlito", "DejaVu Sans", "Noto Sans CJK SC", "Noto Sans", Arial, "Liberation Sans", sans-serif');
        expect(family(".text-run:nth-of-type(3)")).toBe('"Odd \\"Face\\"", "DejaVu Sans", "Noto Sans CJK SC", Arial, "Liberation Sans", sans-serif');
        expect(family(".text-run:nth-of-type(4)")).toBe('"Cambria", "Caladea", "DejaVu Serif", "DejaVu Sans", "Noto Serif CJK SC", Arial, "Liberation Sans", sans-serif');
    });

    test("appends Chromium's fallback list after the run's fallback faces, which keep their order", () => {
        const html = `<html><body><span class="text-run" style="font-family: &quot;Calibri&quot;, &quot;Arial&quot;, &quot;Cambria Math&quot;; font-size: 16px;">x</span></body></html>`;

        expect(printHtml(html)).toContain("font-family: &quot;Calibri&quot;, &quot;Carlito&quot;, &quot;Arial&quot;, &quot;Cambria Math&quot;, &quot;Noto Sans&quot;, Arial, &quot;Liberation Sans&quot;, sans-serif;");
    });

    test("prints U+2011 as U+002D and a word joiner in a run with data-non-breaking-hyphen, word spans included", () => {
        const run = (attrs: string, text: string) => `<html><body><span class="text-run"${attrs} style="font-family: &quot;Arial&quot;;">${text}</span></body></html>`;

        expect(printHtml(run(' data-non-breaking-hyphen="hyphen-minus"', 'a‑b&#8209;c <span style="letter-spacing: 1px;">d&#x2011;e</span>'))).toContain('>a-⁠b-⁠c <span style="letter-spacing: 1px;">d-⁠e</span></span>');
        expect(printHtml(run("", "a‑b"))).toContain(">a‑b</span>");
        // "hyphen": U+2010, the LibreOffice target's.
        expect(printHtml(run(' data-non-breaking-hyphen="hyphen"', "a‑b"))).toContain(">a‐⁠b</span>");
    });

    test("draws a run's strike on a span of its own, in the text colour, and keeps the underline's colour and style on the run (#470)", async () => {
        const presentation = new PuppeteerGenPowerPoint();
        const slide = presentation.addSlide();
        const underline = { style: "dotted", color: "FF0000" } as const;

        // A face without installed metrics: its bands are the font's, which only Chromium's own decorations draw.
        slide.addText(
            [
                { text: "Single‑one", options: { strike: "sngStrike", underline } },
                { text: "Double", options: { strike: "dblStrike", underline } },
                { text: "Plain", options: { strike: "sngStrike" } },
            ],
            { x: 1, y: 1, w: 6, h: 1, fontFace: "Placeholder Sans", color: "0000FF" },
        );
        const [single, double, plain] = (await printedPage(presentation)).querySelectorAll<HTMLElement>(".text-run");

        for (const run of [single, double]) {
            const strike = run?.firstElementChild as HTMLElement | null;

            expect([run?.style.textDecorationLine, run?.style.textDecorationStyle, run?.style.textDecorationColor]).toEqual(["underline", "dotted", "#FF0000"]);
            expect([strike?.style.textDecorationLine, strike?.style.textDecorationColor, strike?.style.textDecorationThickness]).toEqual(["line-through", "", "from-font"]);
        }

        expect((single?.firstElementChild as HTMLElement | null)?.style.textDecorationStyle).toBe("solid");
        expect((double?.firstElementChild as HTMLElement | null)?.style.textDecorationStyle).toBe("double");
        // The hyphen rewrite still reaches the text inside the strike span.
        expect(single?.textContent).toBe("Single-⁠one");
        expect([plain?.style.textDecorationLine, plain?.children.length]).toEqual(["line-through", 0]);
    });

    test("leaves a run whose decorations state their bands in em to prepareForPrint, and Chromium draws none of them", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addText([{ text: "Both", options: { strike: "dblStrike", underline: { style: "dbl" } } }, { text: "Plain" }], { x: 1, y: 1, w: 6, h: 1, fontFace: "Arial" });
        const [both, plain] = (await printedPage(presentation)).querySelectorAll<HTMLElement>(".text-run");

        expect([both?.style.textDecorationLine, both?.dataset.printDecoration, both?.children.length]).toEqual(["none", "underline line-through", 0]);
        expect([plain?.style.textDecorationLine, plain?.dataset.printDecoration]).toEqual(["", undefined]);
    });

    test("draws an underline and a strike at the bands the run states, from its baseline (#396, #489)", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addText([{ text: "Underline", options: { strike: "sngStrike", underline: { style: "sng" } } }], { x: 1, y: 1.07, w: 6, h: 1, fontFace: "Arial", fontSize: 24 });
        const document = await presentation.toHtml();
        const [, thickness = "", offset = ""] = document.html.match(/text-decoration-thickness: ([\d.]+)em; text-underline-offset: ([\d.]+)em;/) ?? [];
        const [, strikeOffset = "", strikeThickness = ""] = document.html.match(/data-strike-band="(-[\d.]+)em ([\d.]+)em"/) ?? [];
        const content = pdfContentStreams(await new PuppeteerPdfRenderer().render(document)).join("\n");
        // The SVG's `scale 0 0 scale x y cm` and its two fills (underline, strike) in CSS px; the text's baseline is its
        // `Tm`'s y, at 1 / PRINT_ZOOM of the SVG's scale.
        const [, scale = "", y = "", underlineTop = "", underlineHeight = "", strikeTop = "", strikeHeight = ""] =
            content.match(/\n([\d.]+) 0 0 \1 [\d.]+ ([\d.]+) cm\n(?:(?!Q\n).*\n)*?[\d.]+ ([\d.]+) [\d.]+ ([\d.]+) re\nf\n[\d.]+ ([\d.]+) [\d.]+ ([\d.]+) re\nf\n/) ?? [];
        const [, text = ""] = content.match(/\n1 0 0 -1 [\d.]+ ([\d.]+) Tm\n/) ?? [];
        const baseline = ((Number(text) * Number(scale)) / PRINT_ZOOM - Number(y)) / Number(scale);
        const size = inchesToPixels(24 / 72);

        // Chromium paints the baseline at a whole print device px, 1 / PRINT_ZOOM CSS px: half of it, and float noise.
        expect(Math.abs(Number(underlineTop) - baseline - Number(offset) * size)).toBeLessThanOrEqual(0.5 / PRINT_ZOOM + 0.01);
        expect(Math.abs(Number(strikeTop) - baseline - Number(strikeOffset) * size)).toBeLessThanOrEqual(0.5 / PRINT_ZOOM + 0.01);
        expect([underlineHeight, strikeHeight].map(value => Number(value).toFixed(3))).toEqual([thickness, strikeThickness].map(value => (Number(value) * size).toFixed(3)));
    }, 60_000);

    test("runs a line-band underline's dash pattern on across runs of one look, and starts it again at a colour change (#397)", async () => {
        // A face no runner has, its post band (Arial's) given through fontMetrics, so the HTML states the band in em and
        // the print step draws the dashes whatever fonts are installed (#391).
        const presentation = new PuppeteerGenPowerPoint(undefined, { fontMetrics: { Lorem: { unitsPerEm: 2048, winAscent: 1854, winDescent: 434, underline: { position: -217, thickness: 150 } } } });

        presentation.addSlide().addText(
            [
                { text: "Large ", options: { fontSize: 24, underline: { style: "dash" } } },
                { text: "small ", options: { fontSize: 12, underline: { style: "dash" } } },
                { text: "other", options: { fontSize: 12, underline: { style: "dash", color: "C00000" } } },
            ],
            { x: 1, y: 1, w: 6, h: 1, fontFace: "Lorem", margin: 0, valign: "top" },
        );
        const content = pdfContentStreams(await new PuppeteerPdfRenderer().render(await presentation.toHtml())).join("\n");
        // Each dash is a subpath from its top-left corner; the text is glyphs, so these are the dashes only.
        const lefts = [...content.matchAll(/\n([\d.]+) [\d.]+ m\n/g)].map(match => Number(match[1]));
        const period = (lefts[1] ?? 0) - (lefts[0] ?? 0);
        const restarts = [];

        for (const [index, left] of lefts.entries()) {
            const previous = lefts[index - 1];

            if (previous !== undefined && Math.abs(left - previous - period) > 0.01) {
                restarts.push(index);
            }
        }

        expect(lefts.length).toBeGreaterThan(10);
        expect(restarts.length).toBe(1);
    }, 60_000);

    test("names the object when a run's decoration geometry is malformed", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addText([{ text: "Run", options: { strike: "sngStrike", underline: { style: "dbl" } } }], { x: 1, y: 1, w: 4, h: 1, fontFace: "Arial", objectName: "Title" });
        const document = await presentation.toHtml();
        const renderer = new PuppeteerPdfRenderer();

        await expect(renderer.render({ ...document, html: document.html.replace(/data-underline="[^"]*"/, 'data-underline="line(0)"') })).rejects.toThrow("Cannot print 'Title': a .text-run has data-underline piece 'line(0)'");
        await expect(renderer.render({ ...document, html: document.html.replace(/data-strike-band="[^"]*"/, 'data-strike-band="1px 2px"') })).rejects.toThrow(
            `Cannot print 'Title': a .text-run has data-strike-band="1px 2px", which is not '<offset>em <thickness>em'`,
        );
    }, 60_000);

    test("names the object when a paragraph has no line box", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addText("Text", { x: 1, y: 1, w: 4, h: 1, objectName: "Title" });
        const html = (await presentation.toHtml()).html.replace(/ data-line-baseline="[^"]*"/, "");

        expect(() => printHtml(html)).toThrow("Cannot print 'Title': its .text-paragraph has no numeric data-line-baseline");
    });

    // Chromium paints an inset box-shadow under the box's content, where the <img> would hide it.
    test("draws a picture's inner shadow on a box laid over the picture, inside its rounded frame", async () => {
        const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
        const presentation = new PuppeteerGenPowerPoint();

        presentation
            .addSlide()
            .addImage({ data: PNG, x: 1, y: 1, w: 2, h: 2, rounding: true, shadow: { type: "inner", blur: 3, offset: 3, angle: 45 } })
            .addShape("ellipse", { x: 4, y: 1, w: 1, h: 1, shadow: { type: "outer", blur: 3, offset: 3, angle: 45 } });
        const html = (await presentation.toHtml()).html;
        const inset = html.match(/box-shadow: (inset [^;]+);/)?.[1];

        expect(inset).toStartWith("inset ");
        const image = new JSDOM(printHtml(html)).window.document.querySelector<HTMLElement>(".slide-image");

        expect(image?.style.boxShadow).toBe("none");
        const overlay = image?.lastElementChild;

        expect(overlay?.previousElementSibling?.querySelector("img")).toBeTruthy();
        expect(overlay?.getAttribute("style")).toBe(`position: absolute; inset: 0; border-radius: inherit; box-shadow: ${inset}; pointer-events: none;`);
        expect(image?.style.borderRadius).toBe("50%");
    });

    // LibreOffice, WASM and PDFKit cast a text box's shadow from everything it draws, under it (#626).
    test("casts a text box's shadow from its fill, or from its glyphs alone where it has no fill", async () => {
        const presentation = new PuppeteerGenPowerPoint();
        const slide = presentation.addSlide();
        const shadow = (): PptxGenJS.ShadowProps => ({ type: "outer", blur: 1, offset: 8, angle: 180, color: "000000", opacity: 1 });

        slide.addText("Solid", { x: 1, y: 1, w: 2, h: 1, fill: { color: "70AD47" }, shadow: shadow() });
        slide.addText("Half", { x: 4, y: 1, w: 2, h: 1, fill: { color: "70AD47", transparency: 50 }, shadow: shadow() });
        slide.addText("None", { x: 1, y: 3, w: 2, h: 1, shadow: shadow() });
        slide.addText([{ text: "Runs", options: { bold: true } }], { x: 4, y: 3, w: 2, h: 1, rotate: 30, shadow: shadow() });
        const printed = [...(await printedPage(presentation)).querySelectorAll<HTMLElement>(".slide-text")];
        const boxes = printed.map(box => ({
            box: box.style.boxShadow === "none" ? "none" : "cast",
            glyphs: [...box.querySelectorAll<HTMLElement>(".text-run")].map(run => run.dataset.textShadow),
        }));
        const half = printed[1]?.style;

        // The half-transparent fill casts its share at half strength, laid inside the box under the fill too.
        expect(half?.boxShadow).toContain("rgba(0, 0, 0, 0.5)");
        expect(half?.backgroundColor).toBe("transparent");
        expect(half?.backgroundImage).toBe("linear-gradient(rgba(112, 173, 71, 0.5), rgba(112, 173, 71, 0.5)), linear-gradient(rgba(0, 0, 0, 0.5), rgba(0, 0, 0, 0.5))");
        // Moved by the shadow's offset, as the last box-shadow states it.
        const offsets = [...(printed[1]?.getAttribute("style") ?? "").matchAll(/box-shadow: (\S+px \S+px) /g)].map(([, offset]) => offset);

        expect(printed[1]?.getAttribute("style")).toContain(`background-position: 0px 0px, ${offsets.at(-1)};`);

        // An opaque fill hides the glyphs' share; a run array casts the box's shadow as a string does.
        expect(boxes).toEqual([
            { box: "cast", glyphs: [undefined] },
            { box: "cast", glyphs: ["shadow"] },
            { box: "none", glyphs: ["shadow"] },
            { box: "none", glyphs: ["shadow"] },
        ]);
        const model = presentation.getModel().slides[0]?.objects ?? [];
        const glyphShadow = (index: number) => {
            const object = model[index];

            return object?.kind === "text" ? object.text.paragraphs[0]?.runs[0]?.shadow : undefined;
        };

        // Under the half-transparent fill the glyph shadow shows through it: half 70AD47, half the black shadow.
        expect(String(glyphShadow(1)?.color)).toBe("#385724");
        // The rotated array box's glyph shadow is cast in slide space, as the box's own.
        const boxShadow = model[3]?.style.shadow;

        expect(boxShadow).toBeDefined();
        expect(glyphShadow(3) ?? "none").toEqual(boxShadow ?? "none");
    });

    // addText rejects an inner shadow (tests/shape.test.ts): hand-written HTML can still write one on a text box or run,
    // where Chromium would draw it differently from the WASM renderer, or outside the glyphs.
    test("fails on an inner shadow on a text box or a run", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addText("Inset", { x: 1, y: 1, w: 2, h: 1, objectName: "Box", shadow: { type: "outer", blur: 3, offset: 3, angle: 45 } });
        const html = (await presentation.toHtml()).html;
        const box = html.replace("box-shadow: ", "box-shadow: inset ");

        expect(box).not.toBe(html);
        expect(() => printHtml(box)).toThrow("Cannot print 'Box': its box-shadow is inset, and the profile draws an inner shadow only on a .slide-image");
        const run = html.replace('<span class="text-run"', '<span class="text-run" data-shadow-inset=""');

        expect(run).not.toBe(html);
        expect(() => printHtml(run)).toThrow("Cannot print 'Box': a .text-run has data-shadow-inset");
    });

    // Hand-written HTML can still mark a shape's shadow inset; drop-shadow() would draw it outside the shape (#198).
    test("fails on an inner shadow on a shape instead of drawing it as an outer one", async () => {
        const presentation = new PuppeteerGen();

        presentation.addSlide().addShape("ellipse", { x: 1, y: 1, w: 1, h: 1, objectName: "Inset", shadow: { type: "outer", blur: 3, offset: 3, angle: 45 } });
        const html = (await presentation.toHtml()).html;
        const inner = html.replace('<svg class="shape-svg"', '<svg class="shape-svg" data-shadow-inset=""');

        expect(inner).not.toBe(html);
        expect(() => printHtml(inner)).toThrow("Cannot print 'Inset': its .shape-svg has data-shadow-inset");
    });

    test("changes only the bytes it rewrites, the same way every time", async () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();

        slide.addText("Text", { x: 1, y: 1, w: 4, h: 1 });
        slide.addShape("line", { x: 6.55, y: 1.6, w: 2.6, h: 0, line: { color: "4472C4", width: 2.25, dashType: "dash", beginArrowType: "oval" } });
        const html = (await presentation.toHtml()).html;
        const printed = printHtml(html);

        expect(printHtml(html)).toBe(printed);
        // Only additions inside start tags: the text and the structure around them are untouched.
        const outsideTags = (markup: string) => markup.replace(/<[^>]*>/g, "<>");

        expect(outsideTags(printed)).toBe(outsideTags(html));
        // Numbers are written as JavaScript prints them, the same text the in-page step gave the CSSOM setters.
        expect(printed).toContain('line-height: 28.799999999999997px; tab-size: 96px; position: relative; top: 9.624999999999996px;"');
        expect(printed).toContain('style="overflow: visible; transform: translate(0.05px, -0.025px);"');
    });

    test("keeps data: URI payloads byte for byte while it skips them in the parse", async () => {
        const presentation = new PuppeteerGen();
        const slide = presentation.addSlide();
        const png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

        slide.addImage({ data: png, x: 1, y: 1, w: 1, h: 1 });
        slide.addImage({ data: png.replace("iVBOR", "iVBOr"), x: 3, y: 1, w: 1, h: 1 });
        // Text that looks like a src attribute is escaped text, not an attribute, and stays as written.
        slide.addText('a src="data:x,y<b>z</b>', { x: 1, y: 3, w: 4, h: 1 });
        const html = (await presentation.toHtml()).html;
        const printed = printHtml(html);
        const uris = (markup: string) => [...markup.matchAll(/src="(data:[^"]*)"/g)].map(match => match[1]);

        expect(uris(printed)).toEqual(uris(html));
        expect(uris(printed)).toHaveLength(2);
        expect(new JSDOM(printed).window.document.querySelector(".text-run")?.textContent).toBe('a src="data:x,y<b>z</b>');
        expect(printed).toContain("position: relative; top:");
    });

    test("passes HTML outside the profile through with only the print zoom added", () => {
        const html = `<!DOCTYPE html><html><head><style>h1 { color: red }</style></head><body><h1 class="title" style="top: 1.5px;">Any HTML</h1><svg><path stroke-dasharray="0 2 4"/></svg></body></html>`;

        expect(printHtml(html)).toBe(html.replace("<html>", `<html style="zoom: ${PRINT_ZOOM};">`));
    });
});
