// Browser-safe WASM text renderer ("@pdflapas/wasm" under the browser/default conditions).
// Must not load node:* modules, puppeteer or jsdom (enforced by tests/wasm.test.ts and scripts/build.ts).
import { assertHtmlPresentationDocument, type HtmlPresentationDocument, type MissingFont, type MissingFontPolicy, type ModelHtmlRequest, missingFontPolicy, type PresentationRenderer, reportMissingFonts } from "@pdflapas/model";
import init, { Renderer } from "./pdf_wasm.js";

export { type MissingFont, MissingFontError, type MissingFontPolicy } from "@pdflapas/model";

/** One entry of the binding's `missingFonts()`: a wasm-owned object, released with `free()`. */
export type MissingFontBinding = {
    readonly face: string;
    readonly used?: string | undefined;
    readonly objects: string[];
    free(): void;
};

/** The wasm-bindgen surface the loader relies on; tests/wasm.test.ts checks the generated class against it. */
// eslint-disable-next-line @typescript-eslint/consistent-type-definitions -- a class implements this contract
export interface TextPdfBinding {
    addFont(data: Uint8Array): number;
    renderHtml(html: string): Uint8Array;
    warnings(): string[];
    missingFonts(): MissingFontBinding[];
}

/** Where to load `pdf_wasm_bg.wasm` from; a subset of wasm-bindgen's `InitInput` without DOM-only types. */
export type WasmSource = URL | string | Response | ArrayBuffer | ArrayBufferView | WebAssembly.Module;

export type WasmRendererOptions = {
    /** TTF/OTF/TTC font files. WASM cannot see system fonts; at least one is required. */
    fonts: readonly Uint8Array[];
    /** The `pdf_wasm_bg.wasm` module: a URL, bytes, a Response or a compiled module. Defaults to the file next to this module. */
    wasm?: WasmSource | Promise<WasmSource>;
    /**
     * What to do when the renderer skips something it cannot draw; each skip is one line of
     * `warnings()`. `"error"` (default) throws and lists them; `"warn"` returns the PDF and leaves
     * the list in `warnings()`.
     */
    onUnsupported?: "error" | "warn";
    /**
     * What to do when a text run asks for a face none of `fonts` provides, neither the face nor its metric-compatible
     * alias (Calibri -> Carlito, Arial -> Liberation Sans, ...): `"warn"` (default), `"throw"` (`MissingFontError`),
     * `"ignore"`, or a function receiving the list. Not part of `onUnsupported`: the text is drawn, in a fallback face.
     */
    onMissingFont?: MissingFontPolicy;
};

export type WasmRendererPoolOptions = Omit<WasmRendererOptions, "wasm"> & {
    /**
     * Worker threads, each with its own renderer and its own copy of `fonts` in WASM memory (about 330 MB of RSS per worker with
     * `loadSystemFonts()`'s default set, less with only the faces the decks use). Default: one per core but one, at most 8.
     */
    workers?: number;
    /**
     * `true`: `toPdf()` posts the deck's model to a worker, which builds the HTML there (#614): about 1 ms less CPU per
     * request on the calling thread, for about 30-50 MB more RSS per worker (each loads the HTML builder). Default
     * `false`: the HTML is built on the calling thread and the workers never load the builder.
     */
    buildHtmlInWorkers?: boolean;
};

/** What createWasmRendererPool() returns (Node/Bun): a renderer whose renders run in worker threads. */
export type WasmRendererPool = PresentationRenderer & {
    /** Builds the HTML from the model in a worker, then renders it there; toPdf() uses it when present (#614). Only with `buildHtmlInWorkers: true`. */
    renderModelHtml?: (request: ModelHtmlRequest) => Promise<Uint8Array>;
    /** What the last render to finish could not draw; with several in flight, read the error of the one you need instead. */
    warnings(): string[];
    /** The font faces the last render to finish asked for that none of the fonts provides. */
    missingFonts(): MissingFont[];
    /** Terminates the workers; queued and in-flight renders reject. */
    close(): Promise<void>;
    [Symbol.asyncDispose](): Promise<void>;
};

export type SystemFontOptions = {
    /** Directories searched recursively. Defaults to the platform's system and user font directories. */
    dirs?: readonly string[];
    /** Font file basenames to load. Defaults to the metric-compatible fallback set and the Office faces they stand in for. */
    match?: RegExp;
};

/**
 * PDF backend in WebAssembly instead of Chromium: renders the PdfLapas HTML profile document
 * `toHtml()` returns (docs/HTML_PROFILE.md).
 */
export class WasmRenderer implements PresentationRenderer {
    readonly #binding: TextPdfBinding;
    readonly #onUnsupported: "error" | "warn";
    readonly #onMissingFont: MissingFontPolicy;

    constructor(binding: TextPdfBinding, onUnsupported: "error" | "warn" = "error", onMissingFont?: MissingFontPolicy) {
        this.#binding = binding;
        this.#onUnsupported = onUnsupported;
        this.#onMissingFont = missingFontPolicy(onMissingFont);
    }

    /** Parses the profile HTML in WebAssembly; out-of-profile or malformed HTML throws, naming slide, object and property. Images must be data URIs. */
    async render(document: HtmlPresentationDocument): Promise<Uint8Array> {
        assertHtmlPresentationDocument(document);
        const pdf = this.#binding.renderHtml(document.html);
        const skipped = this.#binding.warnings();

        if (skipped.length && this.#onUnsupported !== "warn") {
            throw new Error(`The WASM renderer cannot draw everything in this presentation (pass { onUnsupported: "warn" } to render anyway):\n- ${skipped.join("\n- ")}`);
        }

        reportMissingFonts(this.missingFonts(), this.#onMissingFont);

        return pdf;
    }

    /** What the last render could not draw. */
    warnings(): string[] {
        return this.#binding.warnings();
    }

    /** The font faces the last render asked for that none of the renderer's fonts provides. */
    missingFonts(): MissingFont[] {
        const missing: MissingFont[] = [];

        for (const entry of this.#binding.missingFonts()) {
            const { face, used, objects } = entry;

            entry.free();
            missing.push({ face, objects, ...(used ? { usedFace: used } : {}) });
        }

        return missing;
    }
}

export async function createWasmRenderer(options: WasmRendererOptions): Promise<WasmRenderer> {
    if (!options.fonts?.length) {
        throw new Error("createWasmRenderer() requires at least one font: pass { fonts: [Uint8Array, ...] } (Node/Bun: await loadSystemFonts())");
    }

    await init({ module_or_path: options.wasm ?? new URL("./pdf_wasm_bg.wasm", import.meta.url) });
    const binding: TextPdfBinding = new Renderer();

    for (const [index, font] of options.fonts.entries()) {
        if (binding.addFont(font) === 0) {
            throw new Error(`Font #${index} contains no usable font faces`);
        }
    }

    return new WasmRenderer(binding, options.onUnsupported, options.onMissingFont);
}

export async function createWasmRendererPool(_options: WasmRendererPoolOptions): Promise<WasmRendererPool> {
    throw new Error("createWasmRendererPool() starts worker threads and is only available under Node/Bun; in a browser use createWasmRenderer(), in a Web Worker to keep the page responsive");
}

export async function loadSystemFonts(_options?: SystemFontOptions): Promise<Uint8Array[]> {
    throw new Error("loadSystemFonts() reads the filesystem and is only available under Node/Bun; in a browser pass font bytes to createWasmRenderer({ fonts })");
}
