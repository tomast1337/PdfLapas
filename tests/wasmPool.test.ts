import { afterAll, describe, expect, test } from "bun:test";
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { Worker } from "node:worker_threads";
import type { HtmlPresentationDocument, PresentationRenderer } from "@pdflapas/model";
import { PuppeteerGen, PuppeteerGenLibreOffice, PuppeteerGenPowerPoint } from "../packages/pptxgenjs/src/node";
import { PAGE_SIZES } from "../packages/pptxgenjs/src/pageLayouts";
import { createWasmRendererPool as browserCreateWasmRendererPool } from "../packages/wasm/src/browser";
import { createWasmRenderer, createWasmRendererPool, loadSystemFonts, WasmWorkerPool } from "../packages/wasm/src/node";
import type { PoolRequest, PoolWorkerData } from "../packages/wasm/src/nodeWorker";
import { populateParityFixture } from "./visual/fixture";
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
    const htmlPool = await createWasmRendererPool({ ...options, workers: 2, buildHtmlInWorkers: true });
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

    afterAll(() => Promise.all([pool.close(), htmlPool.close()]));

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
            // @ts-expect-error: a string where the option is a boolean, as untyped callers pass it.
            await expect(createWasmRendererPool({ fonts, workers: 1, buildHtmlInWorkers: "yes" })).rejects.toThrow(/buildHtmlInWorkers to be true or false, got yes/);
            await expect(createWasmRendererPool({ fonts: [new Uint8Array(8)], workers: 1 })).rejects.toThrow(/Font #0 contains no usable font faces/);
            await expect(browserCreateWasmRendererPool(options)).rejects.toThrow(/only available under Node\/Bun/);
        });

        test("toPdf() renders the same bytes, warnings and errors as rendering toHtml(), on every deck, with the HTML built on the caller (default) or in a worker (buildHtmlInWorkers)", async () => {
            let inWorkers = 0;
            let withoutComponents = 0;
            const { renderModelHtml } = htmlPool;

            if (renderModelHtml === undefined) {
                throw new Error("a pool with buildHtmlInWorkers has no renderModelHtml()");
            }

            // The opt-in pool, counting the renders whose HTML a worker built.
            const counted: PresentationRenderer = {
                render: document => htmlPool.render(document),
                renderModelHtml: request => {
                    inWorkers++;

                    return renderModelHtml(request);
                },
            };
            const populators: [string, (presentation: PuppeteerGen) => void][] = [["fixture", populateParityFixture]];

            for (const [name, deck] of Object.entries(DECKS)) {
                populators.push([name, presentation => deck.populate(presentation)]);
            }

            for (const Gen of [PuppeteerGenLibreOffice, PuppeteerGenPowerPoint]) {
                for (const [name, populate] of populators) {
                    const before = new Gen(undefined, { renderer: pool });
                    const after = new Gen(undefined, { renderer: counted });
                    const onCaller = new Gen(undefined, { renderer: pool });

                    populate(before);
                    populate(after);
                    populate(onCaller);
                    const expected = await before
                        .toHtml()
                        .then(document => pool.render(document))
                        .catch((error: unknown) => error);
                    const expectedWarnings = pool.warnings();
                    const pdfOnCaller = await onCaller.toPdf().catch((error: unknown) => error);

                    expect([name, pdfOnCaller]).toEqual([name, expected]);
                    expect([name, pool.warnings()]).toEqual([name, expectedWarnings]);
                    const pdf = await after.toPdf().catch((error: unknown) => error);

                    expect([name, pdf]).toEqual([name, expected]);
                    expect([name, htmlPool.warnings()]).toEqual([name, expectedWarnings]);
                    expect([name, after.warnings()]).toEqual([name, before.warnings()]);
                    if (!after.getModel().slides.some(slide => !slide.hidden && slide.objects.some(object => object.kind === "component"))) {
                        withoutComponents++;
                    }
                }
            }

            // A deck with components builds its HTML on this thread; every other deck in a worker.
            expect(withoutComponents).toBeGreaterThan(populators.length);
            expect(inWorkers).toBe(withoutComponents);
        }, 60_000);

        test("toPdf() builds a deck with components on the caller's thread", async () => {
            const calls: string[] = [];
            const spy: PresentationRenderer = {
                render: async () => (calls.push("render"), new Uint8Array()),
                renderModelHtml: async () => (calls.push("renderModelHtml"), new Uint8Array()),
            };
            const plain = new PuppeteerGen(PAGE, { renderer: spy });
            const withComponent = new PuppeteerGen(PAGE, { renderer: spy });

            plain.addSlide().addText("plain", { x: 1, y: 1, w: 2, h: 1 });
            withComponent.registerComponent("test/text", 1, { render: props => String(props.label) });
            withComponent.addSlide().addComponent({ type: "test/text", version: 1, props: { label: "A" } }, { x: 1, y: 1, w: 2, h: 1 });
            await plain.toPdf();
            await withComponent.toPdf();
            expect(calls).toEqual(["renderModelHtml", "render"]);
        });

        test("a default pool's workers never receive a model, so never load the HTML builder; buildHtmlInWorkers sends them one", async () => {
            const workerData: PoolWorkerData = { fonts, wasm: await WebAssembly.compile(await readFile(new URL("../packages/wasm/src/pdf_wasm_bg.wasm", import.meta.url))), onUnsupported: "warn" };
            // Every request a worker receives: nodeWorker imports the module only for a "model" request.
            const posted: string[] = [];
            const spawn = () => {
                const thread = new Worker(new URL("../packages/wasm/src/nodeWorker.ts", import.meta.url), { workerData });
                const post = thread.postMessage.bind(thread);

                thread.postMessage = (request: PoolRequest) => {
                    posted.push(request.kind);
                    post(request);
                };

                return thread;
            };

            expect(pool.renderModelHtml).toBeUndefined();
            for (const buildHtmlInWorkers of [false, true]) {
                const probed = new WasmWorkerPool(spawn, 1, "ignore", buildHtmlInWorkers);
                const presentation = new PuppeteerGen(PAGE, { renderer: probed });

                await probed.started();
                presentation.addSlide().addText("probe", { x: 1, y: 1, w: 2, h: 1 });
                expect(await presentation.toPdf()).toEqual(await single.render(await presentation.toHtml()));
                await probed.close();
            }

            expect(posted).toEqual(["document", "model"]);
        });

        test("renderModelHtml() rejects a module without modelHtml, and keeps rendering", async () => {
            const model = new PuppeteerGen(PAGE).getModel();

            await expect(htmlPool.renderModelHtml?.({ model, images: new Map(), htmlModule: new URL("../packages/model/src/index.ts", import.meta.url).href })).rejects.toThrow(/has no modelHtml\(model, images\) export/);
            expect(await htmlPool.render(first)).toEqual(await single.render(first));
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
