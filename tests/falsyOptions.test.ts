import { describe, expect, test } from "bun:test";
import type { NormalizedSlideObject } from "../packages/model/src/index";
import { PuppeteerGen } from "../packages/pptxgenjs/src/html";

// A falsy `hyperlink`, `align` or `valign` (0, "", false, null; none of them in PptxGenJS's declared types) is unset
// in PptxGenJS 4.0.1 (#456). Evidence: decks built with each value through pptxgenjs 4.0.1, slide1.xml unzipped:
// - `hyperlink` on a run, text box, shape or image: no <a:hlinkClick> and no hyperlink relationship.
// - `align` on a text box or run: <a:pPr> without `algn` (left); on a cell or table: none either.
// - `valign` on a text box: <a:bodyPr anchor="ctr"> (its default); on a cell or table: <a:tcPr> without `anchor` (top).
// - slideNumber `align`: <a:pPr algn="l"/> (genXmlSlideNumber sets `left`); `valign`: <a:bodyPr/> without `anchor`.
const FALSY = [0, "", false, null] as const;
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==";
const BOX = { x: 1, y: 1, w: 4, h: 1, fontFace: "Arial" };

type Slide = ReturnType<PuppeteerGen["addSlide"]>;

function objects(build: (slide: Slide) => void): NormalizedSlideObject[] {
    const deck = new PuppeteerGen();

    build(deck.addSlide());

    return deck.getModel().slides[0]?.objects ?? [];
}

function textOf(object: NormalizedSlideObject | undefined) {
    if (object?.kind !== "text") {
        throw new Error(`expected a text object, got ${object?.kind}`);
    }

    return object.text;
}

function cellTextOf(object: NormalizedSlideObject | undefined) {
    if (object?.kind !== "table") {
        throw new Error(`expected a table, got ${object?.kind}`);
    }

    const cell = object.table.rows[0]?.cells[0];

    if (cell === undefined || !("text" in cell)) {
        throw new Error("expected a cell with text");
    }

    return cell.text;
}

describe("falsy hyperlink, align and valign are unset, as in PptxGenJS 4.0.1 (#456)", () => {
    for (const value of FALSY) {
        // `as never`: the value is outside the declared type on purpose.
        const falsy = value as never;

        test(`hyperlink ${JSON.stringify(value)} makes no link on a run, text box, shape or image`, () => {
            const model = objects(slide => {
                slide.addText([{ text: "Placeholder", options: { hyperlink: falsy } }], BOX);
                slide.addText("Placeholder", { ...BOX, hyperlink: falsy });
                slide.addShape("rect", { x: 1, y: 1, w: 2, h: 1, hyperlink: falsy });
                slide.addImage({ data: PNG, x: 1, y: 1, w: 1, h: 1, hyperlink: falsy });
            });

            expect(model).toHaveLength(4);
            expect(JSON.stringify(model)).not.toContain('"link"');
        });

        test(`align and valign ${JSON.stringify(value)} keep the text box's defaults: left, middle`, () => {
            const [box, run] = objects(slide => {
                slide.addText("Placeholder", { ...BOX, align: falsy, valign: falsy });
                slide.addText([{ text: "Placeholder", options: { align: falsy } }], BOX);
            });

            expect(textOf(box).verticalAlign).toBe("middle");
            expect(textOf(box).paragraphs[0]?.align).toBe("left");
            expect(textOf(run).paragraphs[0]?.align).toBe("left");
        });

        test(`align and valign ${JSON.stringify(value)} on a cell or table keep the cell's defaults: left, top`, () => {
            const [cell, table] = objects(slide => {
                slide.addTable([[{ text: "Cell", options: { align: falsy, valign: falsy } }]], { x: 1, y: 1, w: 4 });
                slide.addTable([[{ text: "Cell" }]], { x: 1, y: 3, w: 4, align: falsy, valign: falsy });
            });

            for (const text of [cellTextOf(cell), cellTextOf(table)]) {
                expect(text.verticalAlign).toBe("top");
                expect(text.paragraphs[0]?.align).toBe("left");
            }
        });

        test(`slideNumber align and valign ${JSON.stringify(value)} keep its defaults: left, top`, () => {
            const [, number] = objects(slide => {
                slide.addText("Placeholder", BOX);
                slide.slideNumber = { x: 1, y: 1, align: falsy, valign: falsy };
            });

            expect(textOf(number).paragraphs[0]?.align).toBe("left");
            expect(textOf(number).verticalAlign).toBe("top");
        });
    }
});
