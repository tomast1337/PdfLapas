import { describe, expect, test } from "bun:test";
import type PptxGenJS from "pptxgenjs";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";

// PdfLapas never modifies what the caller passes (AGENTS.md "Design"): every input below is deeply frozen, so a write
// onto it throws in these strict-mode modules, and one object reused across calls draws as copies of it would.

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/** `value` frozen with everything it reaches; typed arrays cannot be frozen and are checked by value instead. */
function deepFreeze<T>(value: T): T {
    if (typeof value !== "object" || value === null || ArrayBuffer.isView(value) || Object.isFrozen(value)) {
        return value;
    }

    Object.freeze(value);
    for (const key of Reflect.ownKeys(value)) {
        // `as`: an own key of an object, read back from it.
        deepFreeze((value as Record<PropertyKey, unknown>)[key]);
    }

    return value;
}

/** Changes every value `value` reaches in place: strings, numbers, booleans, typed arrays, and the order of arrays. */
function scramble(value: unknown): void {
    if (ArrayBuffer.isView(value)) {
        new Uint8Array(value.buffer, value.byteOffset, value.byteLength).fill(7);

        return;
    }

    if (typeof value !== "object" || value === null) {
        return;
    }

    // `as`: an object's own keys, written back onto it.
    const record = value as Record<string, unknown>;

    for (const key of Object.keys(record)) {
        const field = record[key];

        if (typeof field === "string") {
            record[key] = `${field}-changed`;
        } else if (typeof field === "number") {
            record[key] = field + 1;
        } else if (typeof field === "boolean") {
            record[key] = !field;
        } else {
            scramble(field);
        }
    }

    if (Array.isArray(value)) {
        value.reverse();
    }
}

/** Every input of every call below; the slide master and placeholder fills exercise the master path too. */
function inputs() {
    const shadow = { type: "outer", color: "000000", opacity: 0.5, blur: 2, offset: 3, angle: 45 };
    const runOptions = { bold: true, color: "FF0000", breakLine: true, bullet: { type: "number" }, hyperlink: { url: "https://example.com" } };

    return {
        layout: { name: "WIDE_TEST", width: 13.333, height: 7.5 },
        theme: { headFontFace: "Arial", bodyFontFace: "Arial" },
        section: { title: "Intro" },
        master: {
            title: "FROZEN",
            background: { color: "EEF2FF" },
            margin: [1, 0.5, 1, 0.5],
            slideNumber: { x: 9, y: 5.2, color: "FF0000", fontSize: 10 },
            objects: [
                { rect: { x: 0, y: 0, w: "100%", h: 0.6, fill: { color: "1F3A5F" } } },
                { text: { text: "Header", options: { x: 0.3, y: 0.1, w: 5, h: 0.4, color: "FFFFFF", fontSize: 14 } } },
                { line: { x: 0.3, y: 5, w: 9, h: 0, line: { color: "888888", width: 1 } } },
                { image: { data: PNG, x: 9, y: 0.1, w: 0.4, h: 0.4 } },
                { chart: { type: "bar", data: [{ name: "S", labels: ["A", "B"], values: [1, 2] }], opts: { x: 6, y: 1, w: 3, h: 2 } } },
                { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 1, w: 9, h: 3 }, text: "Prompt" } },
                { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.2, w: 9, h: 0.6 }, text: "Title" } },
            ],
        },
        slide: { masterName: "FROZEN" },
        background: { color: "FFFFEE" },
        slideNumber: { x: 0.3, y: "90%", color: "333333", fontSize: 9 },
        shadow,
        textRuns: [
            { text: "Lorem ipsum\ndolor", options: runOptions },
            { text: "sit amet", options: runOptions },
            { text: "consectetur", options: { italic: true, glow: { size: 4, opacity: 0.3, color: "00FF00" }, outline: { size: 1, color: "0000FF" } } },
            { text: "adipiscing" },
        ],
        textOptions: {
            x: 0.5,
            y: 0.5,
            w: 4,
            h: 1.5,
            fontFace: "Arial",
            fontSize: 14,
            shadow,
            margin: [4, 4, 4, 4],
            bullet: true,
            lineSpacingMultiple: 1.2,
            paraSpaceAfter: 3,
            fill: { color: "FFFFFF", transparency: 40 },
            line: { color: "000000", width: 1, dashType: "dash" },
        },
        stringOptions: { x: 0.5, y: 2.2, w: 4, h: 0.6, fontFace: "Arial", color: "112233", shape: "roundRect", rectRadius: 0.1, fill: { color: "DDDDDD" }, shadow },
        placeholderRuns: [{ text: "Fill", options: { bold: true } }],
        placeholderOptions: { placeholder: "body" },
        titleRuns: [{ text: "Title fill", options: { italic: true } }],
        titleOptions: { placeholder: "title", margin: 2 },
        shapeOptions: { x: 5, y: 0.5, w: 2, h: 1, fill: { color: "70AD47" }, line: { color: "FF0000", width: 2, beginArrowType: "triangle" }, shadow, rotate: 15, flipH: true },
        custGeomOptions: {
            x: 7.5,
            y: 0.5,
            w: 1,
            h: 1,
            fill: { color: "4472C4" },
            points: [{ x: 0, y: 0 }, { x: 1, y: 0 }, { x: 0.5, y: 1, curve: { type: "quadratic", x1: 1, y1: 1 } }, { close: true }],
        },
        imageOptions: { data: PNG, x: 5, y: 2, w: 1, h: 1, sizing: { type: "cover", w: 1, h: 0.5 }, shadow, hyperlink: { url: "https://example.com" }, transparency: 20 },
        tableRows: [
            [
                { text: "Head", options: { bold: true, fill: { color: "CCCCCC" }, border: [{ type: "solid", pt: 1, color: "000000" }, { type: "none" }, { pt: 2 }, { type: "dash", pt: 1, color: "FF0000" }], margin: [2, 2, 2, 2] } },
                {
                    text: [
                        { text: "Rich", options: runOptions },
                        { text: "cell", options: { italic: true } },
                    ],
                    options: { colspan: 2 },
                },
            ],
            [{ text: "A" }, { text: "B", options: { rowspan: 2 } }, "C"],
            ["D", "E"],
        ],
        tableOptions: { x: 0.5, y: 3, w: 6, colW: [1.5, 2, 2.5], rowH: [0.4, 0.3, 0.3], fontFace: "Arial", fontSize: 10, border: { type: "solid", pt: 1, color: "404040" }, margin: [1, 2, 1, 2] },
        autoPageRows: Array.from({ length: 30 }, (_, index) => [{ text: `Row ${index}` }, `Value ${index}`]),
        autoPageOptions: { x: 0.5, y: 1, colW: [2, 2], autoPage: true, autoPageRepeatHeader: true, autoPageHeaderRows: 1, fontFace: "Arial", rowH: 0.3 },
        chartData: [
            { name: "Series 1", labels: ["Q1", "Q2", "Q3"], values: [3, 5, 4] },
            { name: "Series 2", labels: ["Q1", "Q2", "Q3"], values: [2, 1, 6] },
        ],
        chartOptions: { x: 7, y: 3, w: 4, h: 3, chartColors: ["4472C4", "ED7D31"], showLegend: true, legendPos: "b", showValue: true, catAxisLabelColor: "333333", valGridLine: { color: "DDDDDD", style: "dash", size: 1 }, lineDataSymbolSize: 6 },
        bubbleData: [
            { name: "X", values: [1, 2, 3] },
            { name: "Y", values: [4, 5, 6], sizes: [5, 10, 15] },
        ],
        comboTypes: [
            { type: "bar", data: [{ name: "Bars", labels: ["A", "B"], values: [1, 2] }], options: { barGrouping: "clustered", chartColors: ["4472C4"] } },
            { type: "line", data: [{ name: "Line", labels: ["A", "B"], values: [2, 1] }], options: { secondaryValAxis: true, chartColors: ["ED7D31"] } },
        ],
        comboOptions: {
            x: 7,
            y: 0.5,
            w: 4,
            h: 2.5,
            valAxes: [
                { showValAxisTitle: true, valAxisTitle: "Left" },
                { showValAxisTitle: true, valAxisTitle: "Right" },
            ],
        },
        tableToSlidesHtml: "<table><thead><tr><th>H1</th><th>H2</th></tr></thead><tbody><tr><td>1</td><td>2</td></tr></tbody></table>",
        tableToSlidesOptions: {
            x: 0.5,
            y: 1,
            colW: [3, 3],
            autoPageRepeatHeader: true,
            addText: { text: [{ text: "Extra" }], options: { x: 0.5, y: 0.2, w: 3, h: 0.5, shadow } },
            addShape: { shapeName: "rect", options: { x: 9, y: 0.2, w: 1, h: 0.5, shadow } },
        },
        componentInput: { type: "frozen/card", version: 1, props: { label: "Card", items: ["a", "b"], nested: { n: 1 } } },
        componentOptions: { x: 9, y: 5, w: 2, h: 1 },
        writeProps: { outputType: "base64", compression: true },
    };
}

/** Adds every call of the public API with the inputs, returning the presentation. */
function buildDeck(input: ReturnType<typeof inputs>): PuppeteerGen {
    // A renderer that returns fixed bytes, so write() and stream() run without a browser.
    const pres = new PuppeteerGen(undefined, { renderer: { render: async () => new Uint8Array([37, 80, 68, 70]) }, resolveFontMetrics: null, resolveGlyphAdvances: null });

    pres.defineLayout(input.layout);
    pres.layout = input.layout.name;
    pres.theme = input.theme;
    pres.rtlMode = false;
    pres.author = "Author";
    pres.addSection(input.section);
    // `as`: PptxGenJS's declarations type these objects narrower than the frozen literals read.
    pres.defineSlideMaster(input.master as unknown as PptxGenJS.SlideMasterProps);
    pres.registerComponent(input.componentInput.type, input.componentInput.version, { render: props => `<b>${String(props.label)}</b>` });

    const slide = pres.addSlide(input.slide);

    slide.background = input.background;
    slide.slideNumber = input.slideNumber as PptxGenJS.SlideNumberProps;
    slide.addText(input.textRuns as unknown as PptxGenJS.TextProps[], input.textOptions as PptxGenJS.TextPropsOptions);
    slide.addText("String box", input.stringOptions as PptxGenJS.TextPropsOptions);
    slide.addText(input.placeholderRuns as PptxGenJS.TextProps[], input.placeholderOptions);
    slide.addText(input.titleRuns as PptxGenJS.TextProps[], input.titleOptions);
    slide.addShape("rect", input.shapeOptions as PptxGenJS.ShapeProps);
    slide.addShape("custGeom", input.custGeomOptions as unknown as PptxGenJS.ShapeProps);
    slide.addImage(input.imageOptions as PptxGenJS.ImageProps);
    slide.addTable(input.tableRows as unknown as PptxGenJS.TableRow[], input.tableOptions as PptxGenJS.TableProps);
    slide.addChart("bar", input.chartData as unknown as PptxGenJS.OptsChartData[], input.chartOptions as PptxGenJS.IChartOpts);
    slide.addChart("line", input.chartData as unknown as PptxGenJS.OptsChartData[], input.chartOptions as PptxGenJS.IChartOpts);
    slide.addChart("bubble", input.bubbleData as unknown as PptxGenJS.OptsChartData[], input.chartOptions as PptxGenJS.IChartOpts);
    slide.addChart(input.comboTypes as unknown as PptxGenJS.IChartMulti[], input.comboOptions as unknown as PptxGenJS.OptsChartData[]);
    slide.addNotes("Notes");
    slide.addComponent(input.componentInput, input.componentOptions);
    pres.addSlide().addTable(input.autoPageRows as unknown as PptxGenJS.TableRow[], input.autoPageOptions as PptxGenJS.TableProps);
    pres.tableToSlides(input.tableToSlidesHtml, input.tableToSlidesOptions as PptxGenJS.TableToSlidesProps);

    return pres;
}

describe("caller input", () => {
    test("no public call writes onto what the caller passes", async () => {
        const input = deepFreeze(inputs());
        const pres = buildDeck(input);

        expect(pres.getModel().slides.length).toBeGreaterThan(2);
        expect((await pres.toHtml()).html).toContain("Lorem ipsum");
        expect(await pres.write(input.writeProps as PptxGenJS.WriteProps)).toBe("JVBERg==");
        expect(await pres.stream(input.writeProps)).toBeInstanceOf(Uint8Array);
    });

    test("changing the inputs after the calls leaves the deck as it was", async () => {
        const input = inputs();
        const pres = buildDeck(input);
        const model = pres.getModel();
        const html = (await pres.toHtml()).html;

        // The changed theme makes the title fill be drawn again (PuppeteerSlide.toModel), from its own copy.
        scramble(input);
        expect(pres.getModel()).toEqual(model);
        expect((await pres.toHtml()).html).toBe(html);
    });

    test("a typed array given as chart values is left as it was", () => {
        const values = new Float64Array([3, 5, 4]);
        const pres = new PuppeteerGen();

        pres.addSlide().addChart("bar", deepFreeze([{ name: "S", labels: ["A", "B", "C"], values }]) as unknown as PptxGenJS.OptsChartData[], { x: 1, y: 1, w: 4, h: 3 });
        pres.getModel();
        expect(Array.from(values)).toEqual([3, 5, 4]);
    });

    test("one shadow and one options object reused by three calls draw as three copies", () => {
        const shadow: PptxGenJS.ShadowProps = { type: "outer", color: "000000", opacity: 0.5, blur: 2, offset: 3, angle: 45 };
        const text = { x: 0.5, y: 0.5, w: 3, h: 1, fontFace: "Arial", shadow };
        const shape = { x: 4, y: 0.5, w: 2, h: 1, fill: { color: "70AD47" }, line: { color: "FF0000", width: 2 }, shadow };
        const image = { data: PNG, x: 7, y: 0.5, w: 1, h: 1, shadow };
        const table = { x: 0.5, y: 2, colW: 1.5, fontFace: "Arial" };
        const chart = { x: 0.5, y: 3.5, w: 3, h: 2, chartColors: ["4472C4"] };
        const data = [{ name: "S", labels: ["A", "B"], values: [1, 2] }];
        const deck = (reuse: boolean) => {
            const pres = new PuppeteerGen();

            for (let index = 0; index < 3; index += 1) {
                const slide = pres.addSlide();
                const copy = <T extends object>(value: T): T => (reuse ? value : structuredClone(value));

                slide.addText("Lorem", copy(text));
                slide.addShape("rect", copy(shape));
                slide.addImage(copy(image));
                // Two then three columns: PptxGenJS carries the first table's width onto a reused object (#642).
                slide.addTable([(index === 0 ? ["A", "B"] : ["A", "B", "C"]).map(cell => ({ text: cell }))], copy(table));
                slide.addChart(index === 1 ? "line" : "bar", copy(data), copy(chart));
            }

            return pres;
        };
        const reused = deck(true);
        const copied = deck(false);

        expect(reused.getModel()).toEqual(copied.getModel());
        expect(copied.warnings()).toEqual([]);
        // Every object of every slide draws its shadow.
        for (const slide of reused.getModel().slides) {
            for (const object of slide.objects.slice(0, 3)) {
                expect(object.style.shadow?.visible).toBe(true);
            }
        }
    });
});
