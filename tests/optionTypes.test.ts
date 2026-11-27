import { describe, expect, test } from "bun:test";
import type PptxGenJS from "pptxgenjs";
import { PuppeteerGen } from "../packages/pptxgenjs/src/html";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const SERIES = [{ name: "S", labels: ["A", "B"], values: [1, 2] }];

// `any`-free stand-in for the plain JavaScript a caller writes: the API receives whatever the caller passes.
type Call = (slide: PptxGenJS.Slide) => void;
const loose = <T>(value: unknown): T => value as T;

async function html(call: Call): Promise<string> {
    const pres = new PuppeteerGen();

    call(pres.addSlide());

    return (await pres.toHtml()).html;
}

describe("API option types (#674)", () => {
    // Each of these reached a string method, an attribute or a loop and threw a JavaScript-engine TypeError.
    const wrong: readonly (readonly [string, Call, RegExp])[] = [
        ["fontFace", s => s.addText("x", loose({ fontFace: 5 })), /^addText: option `fontFace` must be a string, got 5$/],
        ["margin", s => s.addText("x", loose({ margin: true })), /^addText: option `margin` must be a number or an array, got true$/],
        ["line", s => s.addText("x", loose({ line: 5 })), /option `line` must be an object or a string, got 5/],
        ["shadow", s => s.addShape("rect", loose({ shadow: 5 })), /^addShape: option `shadow` must be an object, got 5$/],
        ["fit", s => s.addText("x", loose({ fit: 5 })), /option `fit` must be a string/],
        ["vert", s => s.addText("x", loose({ vert: 5 })), /option `vert` must be a string/],
        ["textDirection", s => s.addTable([[{ text: "x" }]], loose({ textDirection: 5 })), /^addTable: option `textDirection` must be a string/],
        ["tabStops", s => s.addText("x", loose({ tabStops: 5 })), /option `tabStops` must be an array, got 5/],
        ["tabStops entry", s => s.addText("x", loose({ tabStops: ["x"] })), /option `tabStops\[0\]` must be an object, got "x"/],
        ["missing tabStops entry", s => s.addText("x", loose({ tabStops: [null] })), /option `tabStops\[0\]` must be an object, got null/],
        ["custGeom points", s => s.addShape(loose("custGeom"), loose({ points: "x" })), /^addShape: option `points` must be an array, got "x"$/],
        ["lang", s => s.addText("x", loose({ lang: 5 })), /option `lang` must be a string/],
        ["objectName", s => s.addImage(loose({ data: PNG, objectName: 5 })), /^addImage: option `objectName` must be a string/],
        ["shapeName", s => s.addShape("rect", loose({ shapeName: 5 })), /option `shapeName` must be a string/],
        ["altText", s => s.addImage(loose({ data: PNG, altText: 5 })), /option `altText` must be a string/],
        ["bullet.style", s => s.addText([{ text: "x", options: loose({ bullet: { style: 5 } }) }]), /option `bullet.style` must be a string/],
        ["sizing.type", s => s.addImage(loose({ data: PNG, sizing: { type: 5, w: 1, h: 1 } })), /option `sizing.type` must be a string/],
        ["border entry", s => s.addTable([[{ text: "x", options: loose({ border: [5] }) }]]), /option `border\[0\]` must be an object, got 5/],
        ["valAxes", s => s.addChart("bar", SERIES, loose({ valAxes: 5 })), /^addChart: option `valAxes` must be an array/],
        ["valAxisMinVal", s => s.addChart("bar", SERIES, loose({ valAxisMinVal: "x" })), /option `valAxisMinVal` must be a number, got "x"/],
        ["series labels", s => s.addChart("bar", loose([{ labels: 5, values: [1] }])), /option `labels` must be an array, got 5/],
        ["series values", s => s.addChart("bar", loose([{ labels: ["A"], values: 5 }])), /option `values` must be an array/],
        ["lineDataSymbol", s => s.addChart("line", SERIES, loose({ lineDataSymbol: 5 })), /option `lineDataSymbol` must be a string/],
        ["hyperlink.tooltip", s => s.addText("x", loose({ hyperlink: { url: "https://example.com", tooltip: 5 } })), /option `hyperlink.tooltip` must be a string/],
        ["run options", s => s.addText([{ text: "x", options: loose({ fontFace: [] }) }]), /option `fontFace` must be a string, got an array/],
        ["cell run options", s => s.addTable([[{ text: [{ text: "x", options: loose({ lang: {} }) }] }]]), /option `lang` must be a string, got an object/],
    ];

    for (const [name, call, message] of wrong) {
        test(`refuses a ${name} of the wrong type, naming it`, async () => {
            await expect(html(call)).rejects.toThrow(message);
        });
    }

    // PptxGenJS reads these with truthiness checks, so null is the option left out.
    const absent: readonly (readonly [string, Call, Call])[] = [
        ["line", s => s.addText("x", loose({ line: null })), s => s.addText("x", {})],
        ["margin", s => s.addText("x", loose({ margin: null })), s => s.addText("x", {})],
        ["fontFace", s => s.addText("x", loose({ fontFace: null })), s => s.addText("x", {})],
        ["catLabelFormatCode", s => s.addChart("bar", SERIES, loose({ catLabelFormatCode: null })), s => s.addChart("bar", SERIES, {})],
        ["sizing.type", s => s.addImage(loose({ data: PNG, sizing: { type: null, w: 1, h: 1 } })), s => s.addImage(loose({ data: PNG, sizing: { w: 1, h: 1 } }))],
    ];

    for (const [name, withNull, without] of absent) {
        test(`draws a null ${name} as if it were left out`, async () => {
            expect(await html(withNull)).toBe(await html(without));
        });
    }

    test("draws the missing sides of a short border array as none, as PptxGenJS does", async () => {
        const pres = new PuppeteerGen();

        pres.addSlide().addTable([[{ text: "x", options: { border: loose([{ type: "solid", pt: 2, color: "FF0000" }]) } }]]);

        const table = pres.getModel().slides[0]?.objects[0];
        const borders = table?.kind === "table" ? table.table.rows[0]?.cells[0]?.borders : undefined;

        expect(borders?.[0]?.width).toBeGreaterThan(0);
        expect(borders?.slice(1).map(side => side.width)).toEqual([0, 0, 0]);
    });
});
