import { describe, expect, spyOn, test } from "bun:test";
import { UnsupportedFeatureError } from "../packages/pptxgenjs/src/html";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

function errorFrom(action: () => unknown): unknown {
    try {
        action();
    } catch (error) {
        return error;
    }

    throw new Error("expected the call to throw");
}

describe("PptxGenJS methods without a PDF rendering", () => {
    test("addMedia throws an UnsupportedFeatureError naming the method and the addImage workaround", () => {
        const slide = new PuppeteerGen().addSlide();
        const error = errorFrom(() => slide.addMedia({ type: "video", path: "movie.mp4", x: 1, y: 1, w: 4, h: 3 }));

        expect(error).toBeInstanceOf(UnsupportedFeatureError);
        expect(error).toMatchObject({ name: "UnsupportedFeatureError", code: "UNSUPPORTED_FEATURE", feature: "addMedia" });
        expect((error as Error).message).toBe("addMedia() is not supported by PuppeteerGen yet. A PDF cannot play audio or video; add a poster frame with addImage() and link the media through its `hyperlink` option.");
    });

    // Filling a slide master's placeholder is not drawn yet (#465); the option throws instead of drawing a plain object.
    test("placeholder throws an UnsupportedFeatureError naming the option and leaves the deck unchanged", () => {
        const pres = new PuppeteerGen();

        pres.defineSlideMaster({ title: "MASTER", objects: [{ placeholder: { options: { name: "body", type: "body", x: 1, y: 1, w: 8, h: 3 } } }] });
        const slide = pres.addSlide({ masterName: "MASTER" });
        const text = errorFrom(() => slide.addText("Title", { placeholder: "body" }));

        expect(text).toBeInstanceOf(UnsupportedFeatureError);
        expect(text).toMatchObject({ feature: "addText({ placeholder })" });
        expect((text as Error).message).toBe("addText({ placeholder }) is not supported by PuppeteerGen yet. Filling a slide master's placeholder is not drawn yet (#465). Remove `placeholder` and give the object the placeholder's x, y, w, h and style.");
        expect(errorFrom(() => slide.addImage({ data: PNG, x: 1, y: 1, w: 1, h: 1, placeholder: "body" }))).toMatchObject({ feature: "addImage({ placeholder })" });
        expect(pres.getModel().slides).toHaveLength(1);
        expect(pres.getModel().slides[0]?.objects).toEqual([]);
    });

    // genTableToSlides: the built-in layout's 0.5in margins replace slideMargin.
    test("tableToSlides masterSlideName DEFAULT takes the built-in layout's margins over slideMargin", () => {
        const html = '<table style="width:600px"><tr><td>A</td></tr></table>';
        const tableX = (options: { slideMargin: number; masterSlideName?: string }) => {
            const pres = new PuppeteerGen();

            pres.tableToSlides(html, options);

            return pres.getModel().slides[0]?.objects[0]?.style.geometry.x ?? Number.NaN;
        };

        expect(tableX({ slideMargin: 1, masterSlideName: "DEFAULT" })).toBe(tableX({ slideMargin: 0.5 }));
        expect(tableX({ slideMargin: 1 })).not.toBe(tableX({ slideMargin: 0.5 }));
    });

    // PptxGenJS stores notes for the notes page only; a slide export never draws them.
    test("addNotes is accepted, chainable and draws nothing on the slide", () => {
        const pres = new PuppeteerGen();
        const slide = pres.addSlide();

        expect(slide.addNotes("Speaker notes").addImage({ data: PNG, x: 1, y: 1, w: 1, h: 1 })).toBe(slide);
        expect(pres.getModel().slides[0]?.objects.map(object => object.kind)).toEqual(["image"]);
    });

    // PptxGenJS records sections for PowerPoint's slide list and warns, without throwing, on a missing title.
    test("addSection is accepted, draws nothing and keeps PptxGenJS's warnings", () => {
        const warn = spyOn(console, "warn").mockImplementation(() => {});
        const pres = new PuppeteerGen();

        pres.addSection({ title: "Intro" });
        expect(warn).not.toHaveBeenCalled();
        pres.addSection({} as { title: string });
        expect(warn).toHaveBeenCalledWith("addSection requires a title");
        warn.mockRestore();
        pres.addSlide();
        expect(pres.getModel().slides).toHaveLength(1);
        expect(pres.getModel().slides[0]?.objects).toEqual([]);
    });
});
