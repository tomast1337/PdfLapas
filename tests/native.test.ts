import { describe, expect, test } from "bun:test";
import { resolve } from "node:path";
import { createNativeRenderer, createNativeRendererPool } from "../packages/native/src/bun";
import { createNativeRenderer as createOutsideBun, createNativeRendererPool as createPoolOutsideBun } from "../packages/native/src/index";
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
    test("the Node entry imports and its factories say to use @pdflapas/wasm until node:ffi", async () => {
        const node = await import("../packages/native/src/node");
        const message = "@pdflapas/native needs Bun (bun:ffi); on Node use @pdflapas/wasm (createWasmRenderer / createWasmRendererPool). Node support waits on node:ffi (#612 phase 4).";

        for (const create of [node.createNativeRenderer, node.createNativeRendererPool]) {
            const error: unknown = await create({ fonts: [new Uint8Array(1)] }).catch((thrown: unknown) => thrown);

            expect(error).toBeInstanceOf(node.UnsupportedRuntimeError);
            expect(error).toMatchObject({ runtime: "node", message });
        }
    });

    test("the default entry says what to use instead", async () => {
        for (const create of [createOutsideBun, createPoolOutsideBun]) {
            await expect(create({ fonts: [new Uint8Array(1)] })).rejects.toThrow("@pdflapas/native needs Bun (bun:ffi), which this runtime does not have; use @pdflapas/wasm");
        }
    });
});

describe.skipIf(!canRender)("native backend render", () => {
    test("renders every visual deck and the shared fixture to the same bytes as the WASM renderer, alone and on pool threads", async () => {
        const options = { fonts: systemFonts, onUnsupported: "warn", onMissingFont: "ignore" } as const;

        using native = await createNativeRenderer(options);
        await using pool = await createNativeRendererPool({ ...options, threads: 4 });
        const wasm = await createWasmRenderer(options);
        const decks: [string, PuppeteerGen][] = [["fixture", new PuppeteerGen()]];

        populateParityFixture(decks[0]![1]);
        for (const [name, deck] of Object.entries(DECKS)) {
            decks.push([name, deck.populate(new PuppeteerGen(PAGE))]);
        }

        const different: string[] = [];
        const pooled: Promise<void>[] = [];

        for (const [name, deck] of decks) {
            const document = await deck.toHtml();
            const [fromWasm, fromNative] = [await wasm.render(document), await native.render(document)];

            if (Buffer.compare(fromWasm, fromNative) !== 0 || JSON.stringify([wasm.warnings(), wasm.missingFonts()]) !== JSON.stringify([native.warnings(), native.missingFonts()])) {
                different.push(name);
            }

            // All decks in flight at once, so the threads render them interleaved.
            pooled.push(
                pool.render(document).then(fromPool => {
                    if (Buffer.compare(fromWasm, fromPool) !== 0) {
                        different.push(`${name} (pool)`);
                    }
                }),
            );
        }

        await Promise.all(pooled);
        expect(different).toEqual([]);
    }, 120_000);

    test("the pool fails each document with the WASM renderer's error and goes on", async () => {
        await using pool = await createNativeRendererPool({ fonts: systemFonts, threads: 2 });
        const wasm = await createWasmRenderer({ fonts: systemFonts });
        const pres = new PuppeteerGen(PAGE);

        pres.addSlide().addText("Hi", { x: 1, y: 1, w: 3, h: 1, objectName: "Title" });
        const document = await pres.toHtml();
        const tampered = { ...document, html: document.html.replace('style="left: 96px;', 'style="filter: blur(2px); left: 96px;') };
        const wasmError = await wasm.render(tampered).then(
            () => "",
            (error: Error) => error.message,
        );
        const [failed, rendered] = await Promise.allSettled([pool.render(tampered), pool.render(document)]);

        expect(failed).toMatchObject({ status: "rejected", reason: { message: wasmError } });
        expect(rendered).toMatchObject({ status: "fulfilled", value: await wasm.render(document) });
    });

    test("the pool applies onMissingFont, frees the event loop and rejects what is in flight when closed", async () => {
        const pool = await createNativeRendererPool({ fonts: systemFonts, threads: 1, onMissingFont: "throw" });
        const pres = new PuppeteerGen(PAGE);

        pres.addSlide().addText("Hi", { x: 1, y: 1, w: 3, h: 1, fontFace: "No Such Face", objectName: "Body" });
        await expect(pool.render(await pres.toHtml())).rejects.toThrow("No Such Face");
        expect(pool.missingFonts()).toMatchObject([{ face: "No Such Face", objects: ["slide 1 'Body'"] }]);
        // render() returns before the thread has drawn anything: a timer queued after it fires first.
        const big = DECKS.shapes!.populate(new PuppeteerGen(PAGE));
        const order: string[] = [];
        const rendering = pool.render(await big.toHtml()).then(() => order.push("rendered"));

        await new Promise(resolve => setTimeout(resolve, 0)).then(() => order.push("timer"));
        await rendering;
        expect(order).toEqual(["timer", "rendered"]);
        const inFlight = pool.render(await big.toHtml());

        await pool.close();
        await expect(inFlight).rejects.toThrow("closed before this document rendered");
        await expect(pool.render(await big.toHtml())).rejects.toThrow("This native renderer pool is closed");
        await expect(createNativeRendererPool({ fonts: systemFonts, threads: 0 })).rejects.toThrow("positive integer number of threads");
        await expect(createNativeRendererPool({ fonts: [new Uint8Array([1, 2, 3])] })).rejects.toThrow("Font #0 contains no usable font faces");
    });

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
