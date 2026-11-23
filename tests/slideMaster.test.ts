import { afterEach, beforeEach, describe, expect, spyOn, test } from "bun:test";
import JSZip from "jszip";
import pptxgen from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";
import { inchesToPixels, pixelsToPoints, pointsToPixels } from "@pdflapas/model";
import type { NormalizedSlide, NormalizedSlideObject, NormalizedTextRun } from "../packages/model/src/index";
import { PuppeteerGen, PuppeteerGenPowerPoint } from "../packages/pptxgenjs/src/node";

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

describe("addText and addImage({ placeholder })", () => {
    const PLACEHOLDER_MASTER: PptxGenJS.SlideMasterProps = {
        title: "PH",
        objects: [
            { rect: { x: 0, y: 0, w: "100%", h: 0.2, fill: { color: "1F3A5F" } } },
            { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9, h: 1 }, text: "Title prompt" } },
            { placeholder: { options: { name: "body", type: "body", x: 0.5, y: 1.5, w: 4.3, h: 3.5, fontSize: 20, fontFace: "Arial" }, text: "Body prompt" } },
            { placeholder: { options: { name: "side", type: "body", x: 5, y: 1.5, w: 4, h: 3 } } },
            // A second placeholder under a name: PptxGenJS fills the first.
            { placeholder: { options: { name: "side", type: "title", x: 0, y: 0, w: 1, h: 1 } } },
        ],
    };

    /**
     * The objects `fill` draws on a slide on PLACEHOLDER_MASTER (its rect left out), beside those `plain` draws on a plain
     * slide, whose text has no space before its first paragraph, as a filled placeholder's has none.
     */
    function filledAndPlain(fill: (slide: PptxGenJS.Slide) => void, plain: (slide: PptxGenJS.Slide) => void, theme?: PptxGenJS.ThemeProps) {
        const filled = new PuppeteerGen();
        const explicit = new PuppeteerGen();

        if (theme) {
            filled.theme = theme;
            explicit.theme = theme;
        }

        filled.defineSlideMaster(PLACEHOLDER_MASTER);
        fill(filled.addSlide({ masterName: "PH" }));
        plain(explicit.addSlide());

        const plainObjects: NormalizedSlideObject[] = [];

        for (const object of explicit.getModel().slides[0]?.objects ?? []) {
            const [first, ...later] = object.kind === "text" ? object.text.paragraphs : [];

            plainObjects.push(object.kind === "text" && first ? { ...object, text: { ...object.text, paragraphs: [{ ...first, spaceBefore: 0 }, ...later] } } : object);
        }

        return { filled: filled.getModel().slides[0]?.objects.slice(1) ?? [], plain: plainObjects, warnings: filled.warnings() };
    }

    test("the placeholder's options win over the text's, as PptxGenJS's runtime merges them", () => {
        const reference = new pptxgen();
        const options: PptxGenJS.TextPropsOptions = { placeholder: "body", fontSize: 30, bold: true, x: 0, y: 0, w: 1, h: 1, bullet: true, lineSpacing: 30 };

        reference.defineSlideMaster(PLACEHOLDER_MASTER);
        const slide = reference.addSlide({ masterName: "PH" });

        slide.addText("Body", structuredClone(options));
        // `as`: _slideObjects is PptxGenJS's internal list of the slide's objects, which its declaration leaves out.
        const [written] = (slide as unknown as { _slideObjects: Array<{ options: PptxGenJS.TextPropsOptions }> })._slideObjects;

        expect(written?.options).toMatchObject({ fontSize: 20, bold: true, x: 0.5, y: 1.5, w: 4.3, h: 3.5, bullet: false, lineSpacing: null, objectName: "Text 1" });

        const { filled, plain, warnings } = filledAndPlain(
            target => target.addText("Body", options),
            target => target.addText("Body", { x: 0.5, y: 1.5, w: 4.3, h: 3.5, fontSize: 20, fontFace: "Arial", bold: true, valign: "top", color: "000000", paraSpaceBefore: 4, objectName: "Text 1" }),
        );

        expect(filled).toEqual(plain);
        expect(warnings).toEqual([]);
    });

    test("what neither sets comes from the master's titleStyle, top-anchored and black whatever the slide's color", () => {
        const { filled, plain } = filledAndPlain(
            target => {
                target.color = "FF0000";
                target.addText("Title", { placeholder: "title" });
            },
            target => target.addText("Title", { x: 0.5, y: 0.3, w: 9, h: 1, fontSize: 44, align: "center", valign: "top", fontFace: "Arial Black", color: "000000", paraSpaceBefore: 0, objectName: "Text 1" }),
            { headFontFace: "Arial Black", bodyFontFace: "Arial" },
        );

        expect(filled).toEqual(plain);
    });

    test("title placeholder text takes the head font PptxGenJS resolves when it writes the file, set before or after the text", async () => {
        const runs: PptxGenJS.TextProps[] = [{ text: "Run" }, { text: "Arial run", options: { fontFace: "Arial" } }];
        const reference = new pptxgen();

        reference.defineSlideMaster(PLACEHOLDER_MASTER);
        reference.addSlide({ masterName: "PH" }).addText(structuredClone(runs), { placeholder: "title" });
        reference.theme = { headFontFace: "Georgia" };
        // `as`: write() returns a Buffer for the nodebuffer output type.
        const zip = await JSZip.loadAsync((await reference.write({ outputType: "nodebuffer" })) as Buffer);
        const slideXml = (await zip.file("ppt/slides/slide1.xml")?.async("string")) ?? "";
        const themeXml = (await zip.file("ppt/theme/theme1.xml")?.async("string")) ?? "";

        // The runs name no font but their own, so the title's comes from the master's +mj-lt: the theme's major font.
        expect(slideXml.match(/<a:latin typeface="[^"]*"/g)).toEqual(['<a:latin typeface="Arial"']);
        expect(themeXml).toContain('<a:majorFont><a:latin typeface="Georgia"');

        const before = new PuppeteerGen();
        const after = new PuppeteerGen();

        before.theme = { headFontFace: "Georgia" };
        before.defineSlideMaster(PLACEHOLDER_MASTER);
        before.addSlide({ masterName: "PH" }).addText(runs, { placeholder: "title" });
        after.defineSlideMaster(PLACEHOLDER_MASTER);
        after.addSlide({ masterName: "PH" }).addText(runs, { placeholder: "title" });
        expect(after.getModel().slides[0]?.objects[1]).toMatchObject({ text: { paragraphs: [{ runs: [{ fontFace: "Calibri Light" }, { fontFace: "Arial" }] }] } });
        after.theme = { headFontFace: "Georgia" };

        const { slides } = after.getModel();

        expect(slides).toEqual(before.getModel().slides);
        expect(slides[0]?.objects[1]).toMatchObject({ text: { paragraphs: [{ runs: [{ fontFace: "Georgia" }, { fontFace: "Arial" }] }] } });
    });

    test("a body placeholder without a size takes the bodyStyle's 32pt with 20% space before after its first paragraph; a run array keeps each run's color", () => {
        const runs: PptxGenJS.TextProps[] = [
            { text: "A", options: { breakLine: true } },
            { text: "B", options: { color: "00AA00" } },
        ];
        const { filled, plain } = filledAndPlain(
            target => {
                target.color = "FF0000";
                // cleanOpts gives each run its own color, else the box's, else the slide's.
                target.addText(runs, { placeholder: "side" });
            },
            target => {
                target.color = "FF0000";
                target.addText(runs, { x: 5, y: 1.5, w: 4, h: 3, fontSize: 32, valign: "top", paraSpaceBefore: 6.4, objectName: "Text 1" });
            },
        );
        const [text] = filled;

        expect(filled).toEqual(plain);
        // LibreOffice spaces the first paragraph of a plain text box only (editeng's ULSPACEFIRSTPARA), not a placeholder's.
        expect(text?.kind === "text" ? text.text.paragraphs.map(paragraph => paragraph.spaceBefore) : []).toEqual([0, pointsToPixels(6.4)]);
    });

    test("each body placeholder paragraph PptxGenJS writes no spcBef for takes 20% of its own largest run", async () => {
        const runs: PptxGenJS.TextProps[] = [
            { text: "Small", options: { fontSize: 12, breakLine: true } },
            { text: "Small", options: { fontSize: 12 } },
            { text: "Big", options: { fontSize: 40, breakLine: true } },
            { text: "Zero", options: { fontSize: 12, paraSpaceBefore: 0, breakLine: true } },
            { text: "Own", options: { fontSize: 12, paraSpaceBefore: 3 } },
        ];
        const reference = new pptxgen();

        reference.defineSlideMaster(PLACEHOLDER_MASTER);
        reference.addSlide({ masterName: "PH" }).addText(structuredClone(runs), { placeholder: "side" });
        // `as`: write() returns a Buffer for the nodebuffer output type.
        const zip = await JSZip.loadAsync((await reference.write({ outputType: "nodebuffer" })) as Buffer);
        const slideXml = (await zip.file("ppt/slides/slide1.xml")?.async("string")) ?? "";
        const masterXml = (await zip.file("ppt/slideMasters/slideMaster1.xml")?.async("string")) ?? "";

        // Only the last paragraph writes its own space before; the others take the bodyStyle's percentage.
        expect(slideXml.match(/<a:spcBef>.*?<\/a:spcBef>/g)).toEqual(['<a:spcBef><a:spcPts val="300"/></a:spcBef>']);
        expect(masterXml).toMatch(/<p:bodyStyle>\s*<a:lvl1pPr[^>]*><a:spcBef><a:spcPct val="20000"\/><\/a:spcBef>/);

        const pres = new PuppeteerGen();

        pres.defineSlideMaster(PLACEHOLDER_MASTER);
        pres.addSlide({ masterName: "PH" }).addText(runs, { placeholder: "side" });
        const text = pres.getModel().slides[0]?.objects[1];

        // LibreOffice takes the percentage of the paragraph's largest run, not its first or the box's 32pt.
        expect(text?.kind === "text" ? text.text.paragraphs.map(paragraph => Math.round(pixelsToPoints(paragraph.spaceBefore) * 100) / 100) : []).toEqual([0, 8, 2.4, 3]);

        const powerPoint = new PuppeteerGenPowerPoint();

        powerPoint.defineSlideMaster(PLACEHOLDER_MASTER);
        powerPoint.addSlide({ masterName: "PH" }).addText(runs, { placeholder: "side" });
        const powerPointText = powerPoint.getModel().slides[0]?.objects[1];

        // PowerPoint takes it of the single line, 1.2x the largest run, unrounded; its own 3pt is rounded as before.
        expect(powerPointText?.kind === "text" ? powerPointText.text.paragraphs.map(paragraph => Math.round(pixelsToPoints(paragraph.spaceBefore) * 100) / 100) : []).toEqual([0, 9.6, 2.88, 3]);
    });

    test('PptxGenJS writes a "pic" or "tbl" placeholder untyped, and its text takes the bodyStyle as a body one does', async () => {
        const master: PptxGenJS.SlideMasterProps = {
            title: "TYPES",
            objects: [
                { placeholder: { options: { name: "pic", type: "pic", x: 0.5, y: 0.5, w: 4, h: 2 } } },
                { placeholder: { options: { name: "tbl", type: "tbl", x: 5, y: 0.5, w: 4, h: 2 } } },
                { placeholder: { options: { name: "chart", type: "chart", x: 0.5, y: 3, w: 4, h: 2 } } },
            ],
        };
        const reference = new pptxgen();

        reference.defineSlideMaster(master);
        // `as`: write() returns a Buffer for the nodebuffer output type.
        const zip = await JSZip.loadAsync((await reference.write({ outputType: "nodebuffer" })) as Buffer);
        const layoutXml = (await zip.file("ppt/slideLayouts/slideLayout2.xml")?.async("string")) ?? "";

        // genXmlPlaceholder looks the type up twice in PLACEHOLDER_TYPES, whose keys are "image" and "table".
        expect(layoutXml.match(/<p:ph[^>]*>/g)?.map(tag => /type="(\w+)"/.exec(tag)?.[1] ?? "none")).toEqual(["none", "none", "chart"]);

        const pres = new PuppeteerGen();

        pres.defineSlideMaster(master);
        const slide = pres.addSlide({ masterName: "TYPES" });

        for (const name of ["pic", "tbl", "chart"]) {
            slide.addText("A\nB", { placeholder: name });
        }

        for (const object of pres.getModel().slides[0]?.objects ?? []) {
            expect(object).toMatchObject({ kind: "text", text: { paragraphs: [{ align: "left", runs: [{ fontSize: pointsToPixels(32) }] }, { spaceBefore: pointsToPixels(6.4) }] } });
        }
    });

    test("an unknown name draws from the object's own options, top-anchored as PptxGenJS writes it, and adds a warnings() line", () => {
        const reference = new pptxgen();

        reference.defineSlideMaster(PLACEHOLDER_MASTER);
        expect(() => reference.addSlide({ masterName: "PH" }).addText("Lost", { placeholder: "nope", x: 1, y: 4, w: 3, h: 0.5 })).not.toThrow();
        expect(() => reference.addSlide().addText("Lost", { placeholder: "nope", x: 1, y: 4, w: 3, h: 0.5 })).not.toThrow();

        const { filled, plain, warnings } = filledAndPlain(
            target => {
                target.addText("Lost", { placeholder: "nope", x: 1, y: 4, w: 3, h: 0.5 });
                target.addImage({ data: PNG, placeholder: "nope", x: 1, y: 1, w: 1, h: 1 });
            },
            target => {
                target.addText("Lost", { x: 1, y: 4, w: 3, h: 0.5, valign: "top", color: "000000" });
                target.addImage({ data: PNG, x: 1, y: 1, w: 1, h: 1 });
            },
        );

        expect(filled).toEqual(plain);
        expect(warnings).toEqual([
            'addText({ placeholder: "nope" }): the slide\'s master has no placeholder of that name, so the object is drawn from its own options, as PptxGenJS does. Name a placeholder of the master in addSlide({ masterName }), or remove `placeholder`.',
            'addImage({ placeholder: "nope" }): the slide\'s master has no placeholder of that name, so the object is drawn from its own options, as PptxGenJS does. Name a placeholder of the master in addSlide({ masterName }), or remove `placeholder`.',
        ]);
        const withoutMaster = new PuppeteerGen();

        withoutMaster.addSlide().addText("Lost", { placeholder: "nope", x: 1, y: 4, w: 3, h: 0.5 });
        expect(withoutMaster.warnings()[0]).toContain("the slide was added without a master");
    });

    test("a picture moves to the placeholder's x and y and keeps its own size, as slideObjectToXml writes it", () => {
        const { filled, plain } = filledAndPlain(
            target => target.addImage({ data: PNG, placeholder: "body", x: 0, y: 0, w: 2, h: 1 }),
            target => target.addImage({ data: PNG, x: 0.5, y: 1.5, w: 2, h: 1 }),
        );

        expect(filled).toEqual(plain);
    });

    test("throws on a placeholder without a name, which no text could fill", () => {
        const pres = new PuppeteerGen();

        expect(() => pres.defineSlideMaster({ title: "M", objects: [{ placeholder: {} as { options: PptxGenJS.PlaceholderProps } }] })).toThrow("objects[0].placeholder needs `options: { name, type, x, y, w, h }`");
    });

    test("toHtml() writes the filled placeholder as an ordinary text element, and the master's prompt text nowhere", async () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster(PLACEHOLDER_MASTER);
        pres.addSlide({ masterName: "PH" }).addText("Filled title", { placeholder: "title" });
        const { html } = await pres.toHtml();

        expect(html).toContain(">Filled title<");
        expect(html).not.toContain("Title prompt");
        expect(html).toContain('data-object-name="Text 1"');
    });
});

describe('the built-in "DEFAULT" master', () => {
    const rows: PptxGenJS.TableRow[] = Array.from({ length: 40 }, (_, index) => [{ text: `Row ${index}` }]);

    test("starts with the slideNumber a slide's setter wrote last, as PptxGenJS's Slide constructor copies the DEFAULT layout's", async () => {
        const reference = new pptxgen();
        const pres = new PuppeteerGen();

        for (const deck of [reference, pres]) {
            deck.defineSlideMaster({ title: "M", slideNumber: { x: 2, y: 5 } });
        }

        // A defined master's slideNumber reaches the deck's master, not the DEFAULT layout.
        expect(pres.addSlide({ masterName: "DEFAULT" }).slideNumber).toBe(reference.addSlide({ masterName: "DEFAULT" }).slideNumber);
        const props: PptxGenJS.SlideNumberProps = { x: 9, y: 5.2, color: "FF0000" };

        reference.addSlide().slideNumber = props;
        pres.addSlide().slideNumber = props;
        expect(pres.addSlide({ masterName: "DEFAULT" }).slideNumber).toEqual(reference.addSlide({ masterName: "DEFAULT" }).slideNumber);
        expect(pres.addSlide().slideNumber).toBe(reference.addSlide().slideNumber);
        const slides = pres.getModel().slides;

        expect(slides[0]?.objects).toEqual([]);
        expect(slides[2]?.objects.at(-1)?.style.geometry).toMatchObject({ x: inchesToPixels(9), y: inchesToPixels(5.2) });
        expect(firstRunOf(slides[2], -1).text).toBe("3");
        expect(slides[3]?.objects).toEqual([]);
        // Two number elements: the slide whose setter wrote the props and the DEFAULT slide added after it.
        expect((await pres.toHtml()).html.split('data-object-name="Slide Number"')).toHaveLength(3);
    });

    // PptxGenJS reads `slideMargin` on autoPage tables, which its TableProps declaration leaves out.
    type AutoPageTableProps = PptxGenJS.TableProps & { slideMargin?: number };
    const plainSlideCount = (slideMargin: number) => {
        const pres = new PuppeteerGen();
        const options: AutoPageTableProps = { x: 0.5, y: 0.5, w: 9, autoPage: true, slideMargin };

        pres.addSlide().addTable(rows, options);

        return pres.getModel().slides.length;
    };

    test("takes its 0.5in margin over an autoPage table's slideMargin, and continues on DEFAULT slides, as PptxGenJS does", async () => {
        const options: AutoPageTableProps = { x: 0.5, y: 0.5, w: 9, autoPage: true, slideMargin: 1.5 };
        const reference = new pptxgen();

        reference.addSlide({ masterName: "DEFAULT" }).addTable(structuredClone(rows), structuredClone(options));
        // `as`: write() returns a Buffer for the nodebuffer output type.
        const zip = await JSZip.loadAsync((await reference.write({ outputType: "nodebuffer" })) as Buffer);
        const slideXml = await zip.file("ppt/slides/slide2.xml")?.async("string");
        const rels = await zip.file("ppt/slides/_rels/slide2.xml.rels")?.async("string");
        const referenceY = Number(/<p:xfrm><a:off x="\d+" y="(\d+)"/.exec(slideXml ?? "")?.[1]) / 914400;
        const pres = new PuppeteerGen();

        pres.addSlide({ masterName: "DEFAULT" }).addTable(rows, options);
        const slides = pres.getModel().slides;

        // The DEFAULT master writes the same blank slideLayout1.xml a plain slide uses.
        expect(rels).toContain("slideLayout1.xml");
        // The 0.5in bottom margin fits more rows per page than the table's 1.5in, so it takes fewer slides.
        expect(slides.length).toBe(Object.keys(zip.files).filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name)).length);
        expect(slides.length).toBeLessThan(plainSlideCount(1.5));
        expect(referenceY).toBe(0.5);
        expect(slides[1]?.objects.at(-1)?.style.geometry.y).toBe(inchesToPixels(referenceY));
    });
});
