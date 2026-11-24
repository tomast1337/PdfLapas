import { describe, expect, test } from "bun:test";
import PptxGenJSRuntime from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";
import { pointsToPixels } from "@pdflapas/model";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { type SharedShadows, writeShadow } from "../packages/pptxgenjs/src/normalize/sharedShadow";

const PIXEL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
const EMU_PER_POINT = 12700;

/** The deck of #633: one shadow object given to a text box, a shape and a picture, in that order. */
function addSharedShadow(slide: PptxGenJS.Slide, shadow: PptxGenJS.ShadowProps): void {
    slide.addText("One", { x: 0.5, y: 0.5, w: 2, h: 1, shadow });
    slide.addShape("rect", { x: 3.5, y: 0.5, w: 2, h: 1, fill: { color: "70AD47" }, shadow });
    slide.addImage({ data: PIXEL, x: 6.5, y: 0.5, w: 2, h: 1, shadow });
}

function newShadow(): PptxGenJS.ShadowProps {
    return { type: "outer", color: "000000", blur: 1, offset: 8, angle: 180, opacity: 1 };
}

/** Each `<a:outerShdw>` PptxGenJS 4.0.1 writes on the slide, object by object. */
async function pptxGenJsShadows(): Promise<Array<{ blurRad: number; dist: number; dir: number; alpha: number }>> {
    const pptx = new PptxGenJSRuntime();

    addSharedShadow(pptx.addSlide(), newShadow());
    // Stored, not deflated, so the slide XML can be read from the archive bytes without a zip reader.
    const archive = Buffer.from((await pptx.write({ outputType: "nodebuffer", compression: false })) as Uint8Array).toString("utf8");
    const slide = /<p:sld [\s\S]*?<\/p:sld>/.exec(archive)?.[0] ?? "";
    const shadows = [];

    for (const [, attributes = "", alpha = ""] of slide.matchAll(/<a:outerShdw([^>]*)>\s*<a:srgbClr val="\w+">\s*<a:alpha val="(\d+)"/g)) {
        const attribute = (name: string) => Number(new RegExp(` ${name}="(\\d+)"`).exec(attributes)?.[1]);

        shadows.push({ blurRad: attribute("blurRad"), dist: attribute("dist"), dir: attribute("dir"), alpha: Number(alpha) });
    }

    return shadows;
}

describe("a shadow object shared between objects (#633)", () => {
    test("records what PptxGenJS 4.0.1 writes for the first, second and third object", async () => {
        const xml = await pptxGenJsShadows();
        const shared: SharedShadows = new WeakMap();
        const shadow = newShadow();
        const kinds = [];

        expect(xml).toEqual([
            { blurRad: 12700, dist: 101600, dir: 10800000, alpha: 100000 },
            { blurRad: 161290000, dist: 1290320000, dir: 648000000000, alpha: 10000000000 },
            { blurRad: 2048383000000, dist: 16387064000000, dir: 38880000000000000, alpha: 1000000000000000 },
        ]);
        for (const written of xml) {
            kinds.push(writeShadow(shared, shadow).kind);
            expect(shared.get(shadow)).toMatchObject({ blur: written.blurRad, offset: written.dist, angle: written.dir, opacity: written.alpha });
        }

        expect(kinds).toEqual(["drawn", "outOfRange", "outOfRange"]);
        // The caller's object is never written onto.
        expect(shadow).toEqual(newShadow());
    });

    test("draws the first object's shadow from its XML values and no shadow on the later ones, with a warning", async () => {
        const [first] = await pptxGenJsShadows();
        const pres = new PuppeteerGen();

        addSharedShadow(pres.addSlide(), newShadow());
        const [text, shape, image] = pres.getModel().slides[0]?.objects ?? [];

        expect(text?.style.shadow).toMatchObject({ visible: true, blur: pointsToPixels((first?.blurRad ?? 0) / EMU_PER_POINT) });
        expect(text?.style.shadow?.offsetX).toBeCloseTo(-pointsToPixels((first?.dist ?? 0) / EMU_PER_POINT));
        expect(shape?.style.shadow?.visible).toBe(false);
        expect(image?.style.shadow?.visible).toBe(false);
        expect(pres.warnings().filter(line => line.includes("`shadow` object was given to an earlier object too"))).toHaveLength(2);
    });

    test("leaves a copy per call as it was, without a warning", () => {
        const pres = new PuppeteerGen();
        const slide = pres.addSlide();
        const shadow = newShadow();

        slide.addText("One", { x: 0.5, y: 0.5, w: 2, h: 1, shadow: { ...shadow } });
        slide.addShape("rect", { x: 3.5, y: 0.5, w: 2, h: 1, shadow: { ...shadow } });
        slide.addImage({ data: PIXEL, x: 6.5, y: 0.5, w: 2, h: 1, shadow: { ...shadow } });
        const shadows = (pres.getModel().slides[0]?.objects ?? []).map(object => object.style.shadow?.visible);

        expect(shadows).toEqual([true, true, true]);
        expect(pres.warnings()).toEqual([]);
    });
});
