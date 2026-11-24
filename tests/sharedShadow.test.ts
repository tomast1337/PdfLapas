import { describe, expect, test } from "bun:test";
import PptxGenJSRuntime from "pptxgenjs";
import type PptxGenJS from "pptxgenjs";
import { pointsToPixels } from "@pdflapas/model";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";

const PIXEL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==";
const EMU_PER_POINT = 12700;

function newShadow(): PptxGenJS.ShadowProps {
    return { type: "outer", color: "000000", blur: 1, offset: 8, angle: 180, opacity: 1 };
}

/** The deck of #633: one shadow object given to a text box, a shape and a picture, in that order, or a copy each. */
function addSharedShadow(slide: PptxGenJS.Slide, shadow: PptxGenJS.ShadowProps, copies = false): void {
    const own = () => (copies ? { ...shadow } : shadow);

    slide.addText("One", { x: 0.5, y: 0.5, w: 2, h: 1, shadow: own() });
    slide.addShape("rect", { x: 3.5, y: 0.5, w: 2, h: 1, fill: { color: "70AD47" }, shadow: own() });
    slide.addImage({ data: PIXEL, x: 6.5, y: 0.5, w: 2, h: 1, shadow: own() });
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
    test("PptxGenJS 4.0.1 converts it in place, so its second and third objects write it out of range", async () => {
        expect(await pptxGenJsShadows()).toEqual([
            { blurRad: 12700, dist: 101600, dir: 10800000, alpha: 100000 },
            { blurRad: 161290000, dist: 1290320000, dir: 648000000000, alpha: 10000000000 },
            { blurRad: 2048383000000, dist: 16387064000000, dir: 38880000000000000, alpha: 1000000000000000 },
        ]);
    });

    test("draws every object's shadow as a copy per call draws it, and notes PptxGenJS's difference", async () => {
        const [first] = await pptxGenJsShadows();
        const shared = new PuppeteerGen();
        const copied = new PuppeteerGen();
        const shadow = newShadow();

        addSharedShadow(shared.addSlide(), shadow);
        addSharedShadow(copied.addSlide(), newShadow(), true);

        expect(shared.getModel()).toEqual(copied.getModel());
        for (const object of shared.getModel().slides[0]?.objects ?? []) {
            expect(object.style.shadow).toMatchObject({ visible: true, blur: pointsToPixels((first?.blurRad ?? 0) / EMU_PER_POINT) });
            expect(object.style.shadow?.offsetX).toBeCloseTo(-pointsToPixels((first?.dist ?? 0) / EMU_PER_POINT));
        }

        expect(shadow).toEqual(newShadow());
        expect(shared.warnings().filter(line => line.includes("`shadow` object was given to an earlier object too"))).toHaveLength(2);
        expect(copied.warnings()).toEqual([]);
    });
});
