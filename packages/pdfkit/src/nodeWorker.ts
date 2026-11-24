// A worker thread of createPdfkitRendererPool() (node.ts), Node/Bun only: one PdfkitRenderer with the pool's fonts,
// parsed once, rendering the documents the pool posts one at a time and transferring the PDF bytes back. It never
// builds HTML, so it never loads the HTML builder.
import type { HtmlPresentationDocument, ImagePixelLimits, MissingFont } from "@pdflapas/model";
import { parentPort, workerData } from "node:worker_threads";
import { PdfkitRenderer, type UnsupportedPolicy } from "./browser";

/** What createPdfkitRendererPool() hands each worker. */
export type PoolWorkerData = { readonly fonts: readonly Uint8Array[]; readonly onUnsupported: UnsupportedPolicy; readonly imageLimits: Required<ImagePixelLimits> };

/** What a worker posts back: once whether it started, then one message per document. */
export type PoolWorkerMessage =
    | { readonly kind: "ready" }
    | { readonly kind: "failed-to-start"; readonly error: Error }
    | { readonly kind: "pdf"; readonly pdf: Uint8Array; readonly warnings: string[]; readonly missingFonts: MissingFont[] }
    | { readonly kind: "error"; readonly error: Error; readonly warnings: string[]; readonly missingFonts: MissingFont[] };

const port = parentPort;

if (port === null) {
    throw new Error("nodeWorker runs only as a worker thread that createPdfkitRendererPool() starts");
}

// Set by createPdfkitRendererPool(), the only code that starts this file as a worker.
const { fonts, onUnsupported, imageLimits } = workerData as PoolWorkerData;
let renderer: PdfkitRenderer | undefined;

try {
    // Missing fonts are reported on the pool's side, where the caller's onMissingFont (possibly a function) lives.
    renderer = new PdfkitRenderer({ fonts, onUnsupported, onMissingFont: "ignore", imageLimits });
} catch (error) {
    port.postMessage({ kind: "failed-to-start", error: error instanceof Error ? error : new Error(String(error)) } satisfies PoolWorkerMessage);
}

if (renderer !== undefined) {
    const started = renderer;

    port.on("message", async (document: HtmlPresentationDocument) => {
        try {
            const pdf = await started.render(document);
            // renderHtml() writes the PDF into a fresh Uint8Array of its own, never a view of shared memory.
            const bytes = pdf.buffer as ArrayBuffer;

            port.postMessage({ kind: "pdf", pdf, warnings: started.warnings(), missingFonts: started.missingFonts() } satisfies PoolWorkerMessage, [bytes]);
        } catch (error) {
            port.postMessage({ kind: "error", error: error instanceof Error ? error : new Error(String(error)), warnings: started.warnings(), missingFonts: started.missingFonts() } satisfies PoolWorkerMessage);
        }
    });
    port.postMessage({ kind: "ready" } satisfies PoolWorkerMessage);
}
