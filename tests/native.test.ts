import { describe, expect, test } from "bun:test";
import { resolve } from "node:path";
import { createNativeRenderer } from "../packages/native/src/bun";
import { createNativeRenderer as createOutsideBun } from "../packages/native/src/index";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { PAGE_SIZES } from "../packages/pptxgenjs/src/pageLayouts";
import { createWasmRenderer, loadSystemFonts } from "../packages/wasm/src/node";
import { populateParityFixture } from "./visual/fixture";
import { DECKS } from "./visual/wasm/decks";

const PAGE = PAGE_SIZES.SCREEN_16X9.landscape;
// Both renderers are needed: the native output is checked against the WASM output of the same pdf-core.
const nativeBuilt = await Bun.file(resolve(import.meta.dir, "../packages/native/src/pdflapas_native.linux-x64.so")).exists();
const wasmBuilt = await Bun.file(resolve(import.meta.dir, "../packages/wasm/src/pdf_wasm_bg.wasm")).exists();
const systemFonts = nativeBuilt && wasmBuilt ? await loadSystemFonts().catch(() => []) : [];
const canRender = nativeBuilt && wasmBuilt && systemFonts.length > 0;

if (!canRender) {
    test.skip(`native backend render (${!nativeBuilt ? "packages/native/src/pdflapas_native.linux-x64.so is not built (bun run build:native)" : !wasmBuilt ? "the WASM renderer is not built" : "no system fonts found"})`, () => {});
}

describe("native backend outside Bun", () => {
    test("the default entry says what to use instead", async () => {
        await expect(createOutsideBun({ fonts: [new Uint8Array(1)] })).rejects.toThrow("@pdflapas/native needs Bun's FFI; use @pdflapas/wasm on Node");
    });
});

describe.skipIf(!canRender)("native backend render", () => {
    test("renders every visual deck and the shared fixture to the same bytes as the WASM renderer", async () => {
        using native = await createNativeRenderer({ fonts: systemFonts, onUnsupported: "warn", onMissingFont: "ignore" });
        const wasm = await createWasmRenderer({ fonts: systemFonts, onUnsupported: "warn", onMissingFont: "ignore" });
        const decks: [string, PuppeteerGen][] = [["fixture", new PuppeteerGen()]];

        populateParityFixture(decks[0]![1]);
        for (const [name, deck] of Object.entries(DECKS)) {
            decks.push([name, deck.populate(new PuppeteerGen(PAGE))]);
        }

        const different: string[] = [];

        for (const [name, deck] of decks) {
            const document = await deck.toHtml();
            const [fromWasm, fromNative] = [await wasm.render(document), await native.render(document)];

            if (Buffer.compare(fromWasm, fromNative) !== 0 || JSON.stringify([wasm.warnings(), wasm.missingFonts()]) !== JSON.stringify([native.warnings(), native.missingFonts()])) {
                different.push(name);
            }
        }

        expect(different).toEqual([]);
    }, 120_000);

    test("fails with the WASM renderer's errors", async () => {
        using native = await createNativeRenderer({ fonts: systemFonts });
        const wasm = await createWasmRenderer({ fonts: systemFonts });
        const pres = new PuppeteerGen(PAGE);

        pres.addSlide().addText("Hi", { x: 1, y: 1, w: 3, h: 1, objectName: "Title" });
        const { html, pageSize } = await pres.toHtml();
        const tampered = { html: html.replace('style="left: 96px;', 'style="filter: blur(2px); left: 96px;'), pageSize };
        const wasmError = await wasm.render(tampered).then(
            () => "",
            (error: Error) => error.message,
        );

        expect(wasmError).toContain("slide 1: object 'Title': CSS property 'filter'");
        await expect(native.render(tampered)).rejects.toThrow(wasmError);
        // The renderer stays usable after an error.
        expect(new TextDecoder("latin1").decode(await native.render({ html, pageSize })).startsWith("%PDF-")).toBe(true);
    });

    test("applies onUnsupported and onMissingFont as the WASM renderer does", async () => {
        using native = await createNativeRenderer({ fonts: systemFonts, onMissingFont: "throw" });
        const pres = new PuppeteerGen(PAGE);

        pres.addSlide().addText("Hi", { x: 1, y: 1, w: 3, h: 1, fontFace: "No Such Face", objectName: "Body" });
        await expect(native.render(await pres.toHtml())).rejects.toThrow("No Such Face");
        expect(native.missingFonts()).toMatchObject([{ face: "No Such Face", objects: ["slide 1 'Body'"] }]);
    });

    test("rejects fonts it cannot use and renders nothing once closed", async () => {
        await expect(createNativeRenderer({ fonts: [] })).rejects.toThrow("requires at least one font");
        await expect(createNativeRenderer({ fonts: [new Uint8Array([1, 2, 3])] })).rejects.toThrow("Font #0 contains no usable font faces");
        const native = await createNativeRenderer({ fonts: systemFonts });

        native.close();
        native.close();
        await expect(native.render(await new PuppeteerGen(PAGE).toHtml())).rejects.toThrow("This native renderer is closed");
    });
});
