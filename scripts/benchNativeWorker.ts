// A worker thread of the "native-pool" engine of `bun run bench:load`: packages/wasm/src/nodeWorker.ts with a native
// renderer in place of the WASM one, speaking the same PoolWorkerMessage protocol, so WasmWorkerPool drives it unchanged.
// shortcut: bench-only, to measure native renders across cores; a public native pool is phase 3 (#612).
import { parentPort, workerData } from "node:worker_threads";
import { createNativeRenderer } from "../packages/native/src/bun";
import type { HtmlPresentationDocument } from "../packages/model/src/index";
import type { PoolWorkerMessage } from "../packages/wasm/src/nodeWorker";

const port = parentPort;

if (port === null) {
    throw new Error("benchNativeWorker runs only as a worker thread that scripts/bench-load.ts starts");
}

// Set by scripts/bench-load.ts, the only code that starts this file.
const { fonts } = workerData as { readonly fonts: readonly Uint8Array[] };

try {
    const renderer = await createNativeRenderer({ fonts, onUnsupported: "warn", onMissingFont: "ignore" });

    port.on("message", async (document: HtmlPresentationDocument) => {
        try {
            const pdf = await renderer.render(document);

            port.postMessage({ kind: "pdf", pdf, warnings: renderer.warnings(), missingFonts: renderer.missingFonts() } satisfies PoolWorkerMessage, [pdf.buffer as ArrayBuffer]);
        } catch (error) {
            port.postMessage({ kind: "error", error: error instanceof Error ? error : new Error(String(error)), warnings: renderer.warnings(), missingFonts: renderer.missingFonts() } satisfies PoolWorkerMessage);
        }
    });
    port.postMessage({ kind: "ready" } satisfies PoolWorkerMessage);
} catch (error) {
    port.postMessage({ kind: "failed-to-start", error: error instanceof Error ? error : new Error(String(error)) } satisfies PoolWorkerMessage);
}
