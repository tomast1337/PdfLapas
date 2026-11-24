import { afterAll, describe, expect, test } from "bun:test";
import { Worker } from "node:worker_threads";
import { type HtmlPresentationDocument, IMAGE_PIXEL_LIMITS } from "@pdflapas/model";
import { createPdfkitRendererPool as browserCreatePdfkitRendererPool } from "../packages/pdfkit/src/browser";
import { createPdfkitRenderer, createPdfkitRendererPool, loadSystemFonts, PdfkitWorkerPool } from "../packages/pdfkit/src/node";
import type { PoolWorkerData } from "../packages/pdfkit/src/nodeWorker";
import { PuppeteerGen, PuppeteerGenLibreOffice, PuppeteerGenPowerPoint } from "../packages/pptxgenjs/src/node";
import { PAGE_SIZES } from "../packages/pptxgenjs/src/pageLayouts";
import { populateParityFixture } from "./visual/fixture";
import { DECKS } from "./visual/wasm/decks";

const PAGE = PAGE_SIZES.SCREEN_16X9.landscape;
// The faces the decks below draw with, not the whole default set: every worker parses its own.
const fonts = await loadSystemFonts({ match: /^(?:LiberationSans-.*|LiberationSerif-Regular|DejaVuSans(?:-Bold)?|NotoSansCJK-Regular)\.(?:ttf|ttc|otf)$/i });
const options = { fonts, onUnsupported: "warn", onMissingFont: "ignore" } as const;
const single = await createPdfkitRenderer(options);
const pool = await createPdfkitRendererPool({ ...options, workers: 2 });
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

const [first, , images] = documents;

if (first === undefined || images === undefined) {
    throw new Error("no documents");
}

// An image that is not a data URI, which PdfkitRenderer rejects.
const brokenDocument: HtmlPresentationDocument = { ...images, html: images.html.replace(/src="data:image\/[^"]*"/, 'src="image.png"') };
const workerScript = new URL("../packages/pdfkit/src/nodeWorker.ts", import.meta.url);
const workerData: PoolWorkerData = { fonts, onUnsupported: "warn", imageLimits: IMAGE_PIXEL_LIMITS };

afterAll(() => pool.close());

describe("createPdfkitRendererPool", () => {
    test("renders the same bytes, warnings and errors as one PdfkitRenderer, on every deck and both targets", async () => {
        const populators: [string, (presentation: PuppeteerGen) => void][] = [["fixture", populateParityFixture]];

        for (const [name, deck] of Object.entries(DECKS)) {
            populators.push([name, presentation => deck.populate(presentation)]);
        }

        for (const Gen of [PuppeteerGenLibreOffice, PuppeteerGenPowerPoint]) {
            for (const [name, populate] of populators) {
                const presentation = new Gen();

                populate(presentation);
                const document = await presentation.toHtml();
                const expected = await single.render(document).catch((error: unknown) => error);
                const pdf = await pool.render(document).catch((error: unknown) => error);

                expect([name, pdf]).toEqual([name, expected]);
                expect([name, pool.warnings()]).toEqual([name, single.warnings()]);
                expect([name, pool.missingFonts()]).toEqual([name, single.missingFonts()]);
            }
        }
    }, 120_000);

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
    }, 60_000);

    test("rejects with the error PdfkitRenderer throws, and keeps rendering", async () => {
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

    test("applies onUnsupported and onMissingFont as one PdfkitRenderer does", async () => {
        const chart = new PuppeteerGen(PAGE);

        chart.addSlide().addChart("bar", [{ name: "A", labels: ["x", "y"], values: [1, 2] }], { x: 1, y: 1, w: 4, h: 3 });
        const document = await chart.toHtml();
        const strict = await createPdfkitRendererPool({ fonts, workers: 1, onMissingFont: "ignore" });
        const expected = await (await createPdfkitRenderer({ fonts, onMissingFont: "ignore" })).render(document).catch((error: unknown) => error);

        expect(await strict.render(document).catch((error: unknown) => error)).toEqual(expected);
        expect(strict.warnings()).toEqual(["slide 1: object 'Chart 0': charts are not drawn by the PDFKit renderer"]);
        await strict.close();
        const missing = new PuppeteerGen(PAGE);

        missing.addSlide().addText("Missing", { x: 1, y: 1, w: 4, h: 1, fontFace: "No Such Face" });
        const reported: string[] = [];
        const reporting = await createPdfkitRendererPool({ fonts, workers: 1, onMissingFont: list => reported.push(...list.map(font => font.face)) });

        await reporting.render(await missing.toHtml());
        expect(reported).toEqual(["No Such Face"]);
        await reporting.close();
    });

    test("rejects a render whose worker dies, replaces the worker, and renders the next one", async () => {
        const threads: Worker[] = [];
        const killable = new PdfkitWorkerPool(
            () => {
                const thread = new Worker(workerScript, { workerData });

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
        const closing = await createPdfkitRendererPool({ ...options, workers: 1 });
        const pending = closing.render(first).catch((error: unknown) => error);

        await closing.close();
        expect(await pending).toHaveProperty("message", expect.stringMatching(/closed before this document rendered/));
        await expect(closing.render(first)).rejects.toThrow(/is closed/);
        await expect(createPdfkitRendererPool({ fonts: [] })).rejects.toThrow(/requires at least one font/);
        await expect(createPdfkitRendererPool({ fonts, workers: 0 })).rejects.toThrow(/positive integer number of workers, got 0/);
        // @ts-expect-error: a policy outside the type, as untyped callers pass it.
        await expect(createPdfkitRendererPool({ fonts, workers: 1, onUnsupported: "ignore" })).rejects.toThrow(/onUnsupported must be "error" or "warn", got "ignore"/);
        await expect(createPdfkitRendererPool({ fonts, workers: 1, imageLimits: { maxImagePixels: -1 } })).rejects.toThrow(/imageLimits.maxImagePixels must be a positive integer/);
        await expect(createPdfkitRendererPool({ fonts: [new Uint8Array(8)], workers: 1 })).rejects.toThrow(/Font #0 is not a TTF, OTF or TTC font file/);
        await expect(browserCreatePdfkitRendererPool(options)).rejects.toThrow(/only available under Node\/Bun/);
    });

    test("lets the process exit once idle, without close()", () => {
        const script = `
            import { createPdfkitRendererPool, loadSystemFonts } from ${JSON.stringify(new URL("../packages/pdfkit/src/node.ts", import.meta.url).pathname)};
            const pool = await createPdfkitRendererPool({ fonts: await loadSystemFonts({ match: /^LiberationSans-Regular\\.ttf$/ }), workers: 2, onUnsupported: "warn", onMissingFont: "ignore" });
            const pdf = await pool.render(${JSON.stringify(documents[0])});
            console.log(pdf.length > 0 ? "rendered" : "empty");
        `;
        const child = Bun.spawnSync([process.execPath, "-e", script], { timeout: 30_000 });

        expect(child.stdout.toString().trim()).toBe("rendered");
        expect(child.exitCode).toBe(0);
    }, 30_000);
});
