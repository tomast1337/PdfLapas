import { afterAll, describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { Worker } from "node:worker_threads";
import type { HtmlPresentationDocument } from "@pdflapas/model";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { PAGE_SIZES } from "../packages/pptxgenjs/src/pageLayouts";
import { createWasmRendererPool as browserCreateWasmRendererPool } from "../packages/wasm/src/browser";
import { createWasmRenderer, createWasmRendererPool, loadSystemFonts, WasmWorkerPool } from "../packages/wasm/src/node";
import type { PoolWorkerData } from "../packages/wasm/src/nodeWorker";
import { DECKS } from "./visual/wasm/decks";

const PAGE = PAGE_SIZES.SCREEN_16X9.landscape;
// Needs the binary from `bun run build:wasm`; CI downloads it from its "WASM build" job, so there a missing one fails.
const wasmBuilt = await Bun.file(resolve(import.meta.dir, "../packages/wasm/src/pdf_wasm_bg.wasm")).exists();

if (!wasmBuilt && process.env.CI === "true") {
    throw new Error("WASM pool tests cannot run in CI: packages/wasm/src/pdf_wasm_bg.wasm is not built");
}

if (!wasmBuilt) {
    console.warn("Skipping WASM pool tests: packages/wasm/src/pdf_wasm_bg.wasm is not built (bun run build:wasm)");
} else {
    // The faces the decks below draw with, not the whole default set: every worker holds its own copy.
    const fonts = await loadSystemFonts({ match: /^(?:LiberationSans-.*|DejaVuSans(?:-Bold)?|NotoColorEmoji|NotoSansCJK-Regular)\.(?:ttf|ttc|otf)$/i });
    const options = { fonts, onUnsupported: "warn", onMissingFont: "ignore" } as const;
    const single = await createWasmRenderer(options);
    const pool = await createWasmRendererPool({ ...options, workers: 2 });
    const documents: HtmlPresentationDocument[] = [];

    for (const name of ["text", "tables", "images", "emoji", "shapes"]) {
        const deck = DECKS[name];
        const presentation = new PuppeteerGen(PAGE);

        if (deck === undefined) {
            throw new Error(`No ${name} deck in tests/visual/wasm/decks.ts`);
        }

        deck.populate(presentation);
        documents.push(await presentation.toHtml());
    }

    const [first] = documents;

    if (first === undefined) {
        throw new Error("no documents");
    }

    // HTML outside the profile, which WasmRenderer rejects.
    const brokenDocument: HtmlPresentationDocument = { ...first, html: first.html.replace(/<body[^>]*>[\s\S]*<\/body>/, '<body><section class="pdflapas-slide"><img src="image.png"></section></body>') };

    afterAll(() => pool.close());

    describe("createWasmRendererPool", () => {
        test("renders the same bytes, warnings and missing fonts as one WasmRenderer", async () => {
            for (const document of documents) {
                const expected = await single.render(document);
                const pdf = await pool.render(document);

                expect(pdf).toEqual(expected);
                expect(pool.warnings()).toEqual(single.warnings());
                expect(pool.missingFonts()).toEqual(single.missingFonts());
            }
        });

        test("resolves every one of many concurrent renders, each with its own document's bytes", async () => {
            const expected: Uint8Array[] = [];

            for (const document of documents) {
                expected.push(await single.render(document));
            }

            const pending: Promise<Uint8Array>[] = [];

            for (let round = 0; round < 4; round++) {
                for (const document of documents) {
                    pending.push(pool.render(document));
                }
            }

            const pdfs = await Promise.all(pending);

            for (const [index, pdf] of pdfs.entries()) {
                expect(pdf).toEqual(expected[index % expected.length] ?? new Uint8Array());
            }
        });

        test("rejects with the error WasmRenderer throws, and keeps rendering", async () => {
            const expected = await single.render(brokenDocument).then(
                () => "rendered",
                (error: unknown) => error,
            );

            expect(expected).toBeInstanceOf(Error);
            const error = await pool.render(brokenDocument).catch((caught: unknown) => caught);

            expect(error).toBeInstanceOf(Error);
            expect(error).toHaveProperty("message", (expected as Error).message);
            expect(error).toHaveProperty("name", (expected as Error).name);
            expect(await pool.render(first)).toEqual(await single.render(first));
        });

        test("rejects a render whose worker dies, replaces the worker, and renders the next one", async () => {
            const threads: Worker[] = [];
            const workerData: PoolWorkerData = { fonts, wasm: await WebAssembly.compile(await readFile(new URL("../packages/wasm/src/pdf_wasm_bg.wasm", import.meta.url))), onUnsupported: "warn" };
            const killable = new WasmWorkerPool(
                () => {
                    const thread = new Worker(new URL("../packages/wasm/src/nodeWorker.ts", import.meta.url), { workerData });

                    threads.push(thread);

                    return thread;
                },
                1,
                "ignore",
            );

            await killable.started();
            const pending = killable.render(first).catch((error: unknown) => error);

            await threads[0]?.terminate();
            expect(await pending).toHaveProperty("message", expect.stringMatching(/worker exited .* the pool replaced the worker/));
            expect(threads).toHaveLength(2);
            expect(await killable.render(first)).toEqual(await single.render(first));
            await killable.close();
        });

        test("rejects renders after close() and rejects bad options", async () => {
            const closing = await createWasmRendererPool({ ...options, workers: 1 });
            const pending = closing.render(first).catch((error: unknown) => error);

            await closing.close();
            expect(await pending).toHaveProperty("message", expect.stringMatching(/closed before this document rendered/));
            await expect(closing.render(first)).rejects.toThrow(/is closed/);
            await expect(createWasmRendererPool({ fonts: [] })).rejects.toThrow(/requires at least one font/);
            await expect(createWasmRendererPool({ fonts, workers: 0 })).rejects.toThrow(/positive integer number of workers, got 0/);
            await expect(createWasmRendererPool({ fonts: [new Uint8Array(8)], workers: 1 })).rejects.toThrow(/Font #0 contains no usable font faces/);
            await expect(browserCreateWasmRendererPool(options)).rejects.toThrow(/only available under Node\/Bun/);
        });

        test("lets the process exit once idle, without close()", () => {
            const script = `
            import { createWasmRendererPool, loadSystemFonts } from ${JSON.stringify(new URL("../packages/wasm/src/node.ts", import.meta.url).pathname)};
            const pool = await createWasmRendererPool({ fonts: await loadSystemFonts({ match: /^LiberationSans-Regular\\.ttf$/ }), workers: 2, onUnsupported: "warn", onMissingFont: "ignore" });
            const pdf = await pool.render(${JSON.stringify(documents[0])});
            console.log(pdf.length > 0 ? "rendered" : "empty");
        `;
            const child = Bun.spawnSync([process.execPath, "-e", script], { timeout: 30_000 });

            expect(child.stdout.toString().trim()).toBe("rendered");
            expect(child.exitCode).toBe(0);
        }, 30_000);
    });
}
