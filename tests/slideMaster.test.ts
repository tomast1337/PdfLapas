import { afterEach, beforeEach, describe, expect, spyOn, test } from "bun:test";
import pptxgen from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";
import { inchesToPixels } from "@pdflapas/model";
import type { NormalizedSlide, NormalizedSlideObject, NormalizedTextRun } from "../packages/model/src/index";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/** A report master: a header band, a header text, a footer rule, a logo and a slide number. */
const REPORT_MASTER: PptxGenJS.SlideMasterProps = {
    title: "REPORT",
    background: { color: "EEF2FF" },
    margin: [1, 0.5, 1, 0.5],
    slideNumber: { x: 9, y: 5.2, color: "FF0000", fontSize: 10 },
    objects: [
        { rect: { x: 0, y: 0, w: "100%", h: 0.6, fill: { color: "1F3A5F" } } },
        { text: { text: "Header", options: { x: 0.3, y: 0.1, w: 5, h: 0.4, color: "FFFFFF", fontSize: 14 } } },
        { line: { x: 0.3, y: 5, w: 9, h: 0, line: { color: "888888", width: 1 } } },
        { image: { data: PNG, x: 9, y: 0.1, w: 0.4, h: 0.4 } },
        { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 1, w: 9, h: 3 }, text: "Prompt" } },
    ],
};

/** Each object as `kind:objectName`, in paint order. */
function paintOrder(slide: NormalizedSlide | undefined): string[] {
    return (slide?.objects ?? []).map(object => `${object.kind}:${object.style.objectName}`);
}

function textObjectAt(slide: NormalizedSlide | undefined, index: number): Extract<NormalizedSlideObject, { kind: "text" }> {
    const object = slide?.objects.at(index);

    if (object?.kind !== "text") {
        throw new Error(`object ${index} is not text`);
    }

    return object;
}

function firstRunOf(slide: NormalizedSlide | undefined, index: number): NormalizedTextRun {
    const run = textObjectAt(slide, index).text.paragraphs[0]?.runs[0];

    if (!run) {
        throw new Error(`object ${index} has no run`);
    }

    return run;
}

function textOf(slide: NormalizedSlide | undefined, index: number): string {
    const object = textObjectAt(slide, index);
    let text = "";

    for (const paragraph of object.text.paragraphs) {
        for (const run of paragraph.runs) {
            text += run.text;
        }
    }

    return text;
}

let warn: ReturnType<typeof spyOn>;

beforeEach(() => {
    warn = spyOn(console, "warn").mockImplementation(() => {});
});

afterEach(() => {
    warn.mockRestore();
});

describe("defineSlideMaster", () => {
    test("draws the master's objects in order under the slide's own, then the slide number, as PptxGenJS's layout sits under its slide", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster(REPORT_MASTER);
        pres.addSlide({ masterName: "REPORT" }).addText("Body", { x: 1, y: 2, w: 4, h: 1 });
        const [slide] = pres.getModel().slides;

        // The placeholder draws nothing: a layout's prompt text is not shown on its slides.
        expect(paintOrder(slide)).toEqual(["shape:Shape 0", "text:Text 0", "shape:Shape 1", "image:Image 0", "text:Text 0", "text:Slide Number"]);
        expect(textOf(slide, 1)).toBe("Header");
        expect(textOf(slide, 4)).toBe("Body");
        expect(textOf(slide, 5)).toBe("1");
        expect(slide?.objects[0]?.style.geometry).toMatchObject({ x: 0, y: 0, width: inchesToPixels(10), height: inchesToPixels(0.6) });
        expect(warn).not.toHaveBeenCalled();
    });

    test("draws the same master layer on every slide added on it, and nothing on other slides", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster(REPORT_MASTER);
        pres.addSlide({ masterName: "REPORT" });
        pres.addSlide();
        // PptxGenJS 4.0.1 still takes a master name as the first argument (deprecated since 3.2.0).
        pres.addSlide("REPORT");
        const slides = pres.getModel().slides;

        expect(paintOrder(slides[0])).toEqual(paintOrder(slides[2]));
        expect(slides[1]?.objects).toEqual([]);
        expect(textOf(slides[2], 4)).toBe("3");
    });

    test("a slide shows the master's background until it sets its own, which replaces color and picture alike", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster({ title: "COLOR", background: { color: "EEF2FF" } });
        pres.defineSlideMaster({ title: "PICTURE", background: { data: PNG } });
        pres.addSlide({ masterName: "COLOR" });
        pres.addSlide({ masterName: "COLOR" }).background = { color: "FFEEEE" };
        pres.addSlide({ masterName: "PICTURE" });
        pres.addSlide({ masterName: "PICTURE" }).background = { color: "FFEEEE" };
        pres.addSlide({ masterName: "COLOR" }).bkgd = "FFEEEE";
        const plain = new PuppeteerGen();

        plain.addSlide().background = { color: "EEF2FF" };
        plain.addSlide().background = { color: "FFEEEE" };
        const [master, own, picture, ownOverPicture, deprecated] = pres.getModel().slides;

        expect(String(master?.background)).toBe(String(plain.getModel().slides[0]?.background));
        expect(String(own?.background)).toBe(String(plain.getModel().slides[1]?.background));
        expect(paintOrder(picture)).toEqual(["image:Background"]);
        expect(ownOverPicture?.objects).toEqual([]);
        expect(String(ownOverPicture?.background)).toBe(String(own?.background));
        expect(String(deprecated?.background)).toBe(String(own?.background));
    });

    test("the deprecated `bkgd` sets the master's background as PptxGenJS's createSlideMaster does", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster({ title: "OLD", bkgd: "EEF2FF" });
        pres.defineSlideMaster({ title: "NEW", background: { color: "EEF2FF" } });

        pres.addSlide({ masterName: "OLD" });
        pres.addSlide({ masterName: "NEW" });
        pres.addSlide();
        const [old, current, plain] = pres.getModel().slides;

        expect(String(old?.background)).toBe(String(current?.background));
        expect(String(old?.background)).not.toBe(String(plain?.background));
    });

    test("slides added on the master start with its slideNumber, as PptxGenJS's Slide constructor copies the layout's", () => {
        const reference = new pptxgen();
        const pres = new PuppeteerGen();

        reference.defineSlideMaster(REPORT_MASTER);
        pres.defineSlideMaster(REPORT_MASTER);
        expect(pres.addSlide({ masterName: "REPORT" }).slideNumber).toEqual(reference.addSlide({ masterName: "REPORT" }).slideNumber);
        expect(pres.addSlide().slideNumber).toBe(reference.addSlide().slideNumber);
        const own = pres.addSlide({ masterName: "REPORT" });

        own.slideNumber = { x: 1, y: 5 };
        const slides = pres.getModel().slides;

        expect(slides[0]?.objects.at(-1)?.style.geometry).toMatchObject({ x: inchesToPixels(9), y: inchesToPixels(5.2) });
        expect(slides[1]?.objects).toEqual([]);
        // A slide's own number inherits what it leaves out (here color and size) from the master's.
        expect(slides[2]?.objects.at(-1)?.style.geometry).toMatchObject({ x: inchesToPixels(1), y: inchesToPixels(5) });
        expect(firstRunOf(slides[2], -1)).toMatchObject({ text: "3", color: firstRunOf(slides[0], -1).color, fontSize: firstRunOf(slides[0], -1).fontSize });
    });

    test("an unknown master name adds a plain slide, as PptxGenJS does, and warns", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster(REPORT_MASTER);
        pres.addSlide({ masterName: "REPROT" });
        expect(pres.getModel().slides[0]?.objects).toEqual([]);
        expect(warn).toHaveBeenCalledTimes(1);
        expect(String(warn.mock.calls[0]?.[0])).toContain('no slide master is named "REPROT"');
        // PptxGenJS's built-in "DEFAULT" and empty names add a plain slide without a warning.
        pres.addSlide({ masterName: "DEFAULT" });
        pres.addSlide({ masterName: "" });
        expect(pres.getModel().slides).toHaveLength(3);
        expect(warn).toHaveBeenCalledTimes(1);
    });

    test("the first master defined under a title wins, and later changes to the props do not reach it", () => {
        const pres = new PuppeteerGen();
        const props: PptxGenJS.SlideMasterProps = { title: "M", objects: [{ text: { text: "First", options: { x: 1, y: 1, w: 2, h: 1 } } }] };

        pres.defineSlideMaster(props);
        pres.defineSlideMaster({ title: "M", objects: [{ text: { text: "Second", options: { x: 1, y: 1, w: 2, h: 1 } } }] });
        props.objects = [];
        pres.addSlide({ masterName: "M" });

        expect(textOf(pres.getModel().slides[0], 0)).toBe("First");
    });

    test("throws without a title, as PptxGenJS does, and on a margin that is not inches", () => {
        const pres = new PuppeteerGen();

        expect(() => pres.defineSlideMaster({} as PptxGenJS.SlideMasterProps)).toThrow("defineSlideMaster() object argument requires a `title` value");
        expect(() => new pptxgen().defineSlideMaster({} as PptxGenJS.SlideMasterProps)).toThrow("defineSlideMaster() object argument requires a `title` value");
        expect(() => pres.defineSlideMaster({ title: "M", margin: "1in" as unknown as number })).toThrow('`margin` must be inches, a number or [top, right, bottom, left], got "1in"');
    });

    test("a master's text resolves its default font when it is defined, so a later theme font throws as after slide text", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster({ title: "M", objects: [{ text: { text: "Header", options: { x: 1, y: 1, w: 2, h: 1 } } }] });
        expect(() => {
            pres.theme = { bodyFontFace: "Arial" };
        }).toThrow("was set after slide objects were added");
    });

    test("warns about an object key PptxGenJS ignores", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster({ title: "M", objects: [{ table: {} } as unknown as { rect: PptxGenJS.ShapeProps }] });
        pres.addSlide({ masterName: "M" });
        expect(pres.getModel().slides[0]?.objects).toEqual([]);
        expect(String(warn.mock.calls[0]?.[0])).toBe('defineSlideMaster: objects[0] has key "table", which PptxGenJS ignores, so it draws nothing. Use chart, image, line, rect, text or placeholder.');
    });

    test("draws a master chart from its { type, data, opts }", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster({
            title: "M",
            objects: [{ chart: { type: "bar", data: [{ name: "A", labels: ["x"], values: [1] }], opts: { x: 1, y: 1, w: 3, h: 2 } } as unknown as PptxGenJS.IChartOpts }],
        });
        pres.addSlide({ masterName: "M" });

        expect(paintOrder(pres.getModel().slides[0])).toEqual(["chart:Chart 0"]);
    });

    test("toHtml() writes the master's objects as ordinary slide elements before the slide's own", async () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster(REPORT_MASTER);
        pres.addSlide({ masterName: "REPORT" }).addText("Body", { x: 1, y: 2, w: 4, h: 1 });
        const { html } = await pres.toHtml();

        expect(html.indexOf(">Header<")).toBeGreaterThan(-1);
        expect(html.indexOf(">Header<")).toBeLessThan(html.indexOf(">Body<"));
        expect(html).toContain('data-object-name="Slide Number"');
        expect(html).not.toContain("Prompt");
    });
});

describe("a master's margin", () => {
    const rows: PptxGenJS.TableRow[] = Array.from({ length: 40 }, (_, index) => [{ text: `Row ${index}` }]);

    test("starts autoPage continuation tables at its top margin, on slides added on the same master", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster({ title: "M", margin: [1.5, 0.5, 1, 0.5], objects: [{ text: { text: "Header", options: { x: 0, y: 0, w: 5, h: 0.5 } } }] });
        pres.addSlide({ masterName: "M" }).addTable(rows, { x: 0.5, y: 1.5, w: 9, autoPage: true });
        const slides = pres.getModel().slides;

        expect(slides.length).toBeGreaterThan(1);
        expect(paintOrder(slides[1])[0]).toBe("text:Text 0");
        expect(slides[1]?.objects.at(-1)?.style.geometry.y).toBe(inchesToPixels(1.5));
    });

    test("replaces tableToSlides' slideMargin, and every page is added on the master", () => {
        const html = '<table style="width:600px"><tr><td>A</td></tr></table>';
        const tableX = (masterSlideName?: string) => {
            const pres = new PuppeteerGen();

            pres.defineSlideMaster({ title: "M", margin: 1, objects: [{ rect: { x: 0, y: 0, w: 1, h: 1 } }] });
            pres.tableToSlides(html, { slideMargin: 0.25, ...(masterSlideName ? { masterSlideName } : {}) });
            const [slide] = pres.getModel().slides;

            return { order: paintOrder(slide), x: slide?.objects.at(-1)?.style.geometry.x };
        };
        const withMargin = (slideMargin: number) => {
            const pres = new PuppeteerGen();

            pres.tableToSlides(html, { slideMargin });

            return pres.getModel().slides[0]?.objects[0]?.style.geometry.x;
        };

        expect(tableX("M")).toEqual({ order: ["shape:Shape 0", "table:Table 0"], x: withMargin(1) });
        expect(Number(tableX().x)).toBe(Number(withMargin(0.25)));
    });
});
