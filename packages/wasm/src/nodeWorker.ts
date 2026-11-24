// A worker thread of createWasmRendererPool() (node.ts), Node/Bun only: one WasmRenderer with the pool's fonts, loaded
// once, rendering the documents the pool posts one at a time and transferring the PDF bytes back.
import type { DataUri, HtmlPresentationDocument, MissingFont, NormalizedPresentation } from "@pdflapas/model";
import { parentPort, workerData } from "node:worker_threads";
import { createWasmRenderer, type WasmRenderer } from "./browser";

/** What createWasmRendererPool() hands each worker. */
export type PoolWorkerData = { readonly fonts: readonly Uint8Array[]; readonly wasm: WebAssembly.Module; readonly onUnsupported: "error" | "warn" };

/** What the pool posts a worker per render: the HTML document, or the model and the module that builds its HTML (renderModelHtml()). */
export type PoolRequest = { readonly kind: "document"; readonly document: HtmlPresentationDocument } | { readonly kind: "model"; readonly model: NormalizedPresentation; readonly images: ReadonlyMap<string, DataUri>; readonly htmlModule: string };

/** What a worker posts back: once whether it started, then one message per document. */
export type PoolWorkerMessage =
    | { readonly kind: "ready" }
    | { readonly kind: "failed-to-start"; readonly error: Error }
    | { readonly kind: "pdf"; readonly pdf: Uint8Array; readonly warnings: string[]; readonly missingFonts: MissingFont[] }
    | { readonly kind: "error"; readonly error: Error; readonly warnings: string[]; readonly missingFonts: MissingFont[] };

const port = parentPort;

if (port === null) {
    throw new Error("nodeWorker runs only as a worker thread that createWasmRendererPool() starts");
}

// Set by createWasmRendererPool(), the only code that starts this file as a worker.
const { fonts, wasm, onUnsupported } = workerData as PoolWorkerData;
let renderer: WasmRenderer | undefined;

try {
    // Missing fonts are reported on the pool's side, where the caller's onMissingFont (possibly a function) lives.
    renderer = await createWasmRenderer({ fonts, wasm, onUnsupported, onMissingFont: "ignore" });
} catch (error) {
    port.postMessage({ kind: "failed-to-start", error: error instanceof Error ? error : new Error(String(error)) } satisfies PoolWorkerMessage);
}

if (renderer !== undefined) {
    const started = renderer;

    port.on("message", async (request: PoolRequest) => {
        try {
            let document: HtmlPresentationDocument;

            if (request.kind === "document") {
                document = request.document;
            } else {
                // The runtime caches the module after the first import, so it loads once per worker.
                const module: { readonly modelHtml?: unknown } = await import(request.htmlModule);

                if (typeof module.modelHtml !== "function") {
                    throw new TypeError(`renderModelHtml(): ${request.htmlModule} has no modelHtml(model, images) export; name the module @pdflapas/pptxgenjs's Node entry ships`);
                }

                // `as`: the export's contract (ModelHtmlRequest.htmlModule); render() checks the result.
                document = (await module.modelHtml(request.model, request.images)) as HtmlPresentationDocument;
            }

            const pdf = await started.render(document);
            // The binding copies the PDF out of wasm memory into a fresh ArrayBuffer of its own, never a shared one.
            const bytes = pdf.buffer as ArrayBuffer;

            port.postMessage({ kind: "pdf", pdf, warnings: started.warnings(), missingFonts: started.missingFonts() } satisfies PoolWorkerMessage, [bytes]);
        } catch (error) {
            port.postMessage({ kind: "error", error: error instanceof Error ? error : new Error(String(error)), warnings: started.warnings(), missingFonts: started.missingFonts() } satisfies PoolWorkerMessage);
        }
    });
    port.postMessage({ kind: "ready" } satisfies PoolWorkerMessage);
}
