// "@pdflapas/native" under Bun: the Rust renderer (crates/pdf-ffi, pdf-core) as a native library loaded with bun:ffi.
// Same surface, options and errors as @pdflapas/wasm's WasmRenderer; each call into the library is synchronous.
import { assertHtmlPresentationDocument, type HtmlPresentationDocument, type MissingFont, type MissingFontPolicy, missingFontPolicy, reportMissingFonts } from "@pdflapas/model";
import { dlopen, FFIType, type Library, type Pointer, suffix, toArrayBuffer } from "bun:ffi";
import { fileURLToPath } from "node:url";
import type { NativeRenderer, NativeRendererOptions } from "./index";

export type { MissingFont, MissingFontPolicy, NativeRenderer, NativeRendererOptions } from "./index";
export { MissingFontError } from "./index";

/** The C ABI of crates/pdf-ffi/src/lib.rs. `usize` is `u64`: phase 1 builds 64-bit Linux only. */
const SYMBOLS = {
    pdflapas_renderer_new: { args: [], returns: FFIType.ptr },
    pdflapas_add_font: { args: [FFIType.ptr, FFIType.ptr, FFIType.u64], returns: FFIType.u32 },
    pdflapas_render: { args: [FFIType.ptr, FFIType.ptr, FFIType.u64, FFIType.ptr], returns: FFIType.i32 },
    pdflapas_report: { args: [FFIType.ptr, FFIType.ptr], returns: FFIType.i32 },
    pdflapas_last_error: { args: [FFIType.ptr, FFIType.ptr], returns: FFIType.i32 },
    pdflapas_buffer_free: { args: [FFIType.ptr, FFIType.u64], returns: FFIType.void },
    pdflapas_renderer_free: { args: [FFIType.ptr], returns: FFIType.void },
} as const;
/** `PDFLAPAS_OK` in crates/pdf-ffi. */
const OK = 0;
const ENCODER = new TextEncoder();
const DECODER = new TextDecoder();

/** `pdflapas_report`'s JSON: `warnings()` and `missingFonts()` of crates/pdf-wasm. */
type Report = { readonly warnings: string[]; readonly missingFonts: readonly { readonly face: string; readonly used?: string; readonly objects: string[] }[] };

class BunNativeRenderer implements NativeRenderer {
    readonly #library: Library<typeof SYMBOLS>;
    readonly #onUnsupported: "error" | "warn";
    readonly #onMissingFont: MissingFontPolicy;
    /** The `PdflapasBuffer` (`ptr`, `len`) every call writes its output to, reused. */
    readonly #out = new BigUint64Array(2);
    #handle: Pointer | null;
    #warnings: string[] = [];
    #missingFonts: MissingFont[] = [];

    constructor(library: Library<typeof SYMBOLS>, handle: Pointer, onUnsupported: "error" | "warn" = "error", onMissingFont?: MissingFontPolicy) {
        this.#library = library;
        this.#handle = handle;
        this.#onUnsupported = onUnsupported;
        this.#onMissingFont = missingFontPolicy(onMissingFont);
    }

    /** Parses the profile HTML in Rust; out-of-profile or malformed HTML throws, naming slide, object and property. Images must be data URIs. */
    async render(document: HtmlPresentationDocument): Promise<Uint8Array> {
        assertHtmlPresentationDocument(document);
        const handle = this.#open();
        const html = ENCODER.encode(document.html);
        const { symbols } = this.#library;

        if (symbols.pdflapas_render(handle, html, html.byteLength, this.#out) !== OK) {
            symbols.pdflapas_last_error(handle, this.#out);
            throw new Error(DECODER.decode(this.#take()));
        }

        const pdf = this.#take();

        if (symbols.pdflapas_report(handle, this.#out) !== OK) {
            symbols.pdflapas_last_error(handle, this.#out);
            throw new Error(DECODER.decode(this.#take()));
        }

        // The library's own output, built by crates/pdf-ffi's report_json.
        const report = JSON.parse(DECODER.decode(this.#take())) as Report;

        this.#warnings = report.warnings;
        this.#missingFonts = [];
        for (const { face, used, objects } of report.missingFonts) {
            this.#missingFonts.push({ face, objects, ...(used ? { usedFace: used } : {}) });
        }

        if (this.#warnings.length && this.#onUnsupported !== "warn") {
            throw new Error(`The native renderer cannot draw everything in this presentation (pass { onUnsupported: "warn" } to render anyway):\n- ${this.#warnings.join("\n- ")}`);
        }

        reportMissingFonts(this.#missingFonts, this.#onMissingFont);

        return pdf;
    }

    warnings(): string[] {
        return [...this.#warnings];
    }

    missingFonts(): MissingFont[] {
        return [...this.#missingFonts];
    }

    close(): void {
        if (this.#handle !== null) {
            this.#library.symbols.pdflapas_renderer_free(this.#handle);
            this.#handle = null;
            this.#library.close();
        }
    }

    [Symbol.dispose](): void {
        this.close();
    }

    #open(): Pointer {
        if (this.#handle === null) {
            throw new Error("This native renderer is closed; create another with createNativeRenderer()");
        }

        return this.#handle;
    }

    /** Copies the buffer the last call wrote to `#out` into JS memory and frees it in the library. */
    #take(): Uint8Array {
        const [address = 0n, length = 0n] = this.#out;
        // A Rust pointer, never 0 for a written buffer; Number keeps every user-space x86-64 address exactly.
        const pointer = Number(address) as Pointer;
        const bytes = length === 0n ? new Uint8Array() : new Uint8Array(toArrayBuffer(pointer, 0, Number(length))).slice();

        this.#library.symbols.pdflapas_buffer_free(pointer, length);
        this.#out.fill(0n);

        return bytes;
    }
}

/** Loads the platform's library next to this module and the fonts into one renderer. Bun only. */
export async function createNativeRenderer(options: NativeRendererOptions): Promise<NativeRenderer> {
    if (!options.fonts?.length) {
        throw new Error("createNativeRenderer() requires at least one font: pass { fonts: [Uint8Array, ...] } (@pdflapas/wasm's loadSystemFonts() reads the system's)");
    }

    // shortcut: built for linux-x64 only (bun run build:native); per-platform packages are phase 2 (#612).
    const path = fileURLToPath(new URL(`./pdflapas_native.${process.platform}-${process.arch}.${suffix}`, import.meta.url));
    let library: Library<typeof SYMBOLS>;

    try {
        library = dlopen(path, SYMBOLS);
    } catch (error) {
        throw new Error(`No @pdflapas/native library for ${process.platform}-${process.arch} at ${path}: build it with \`bun run build:native\` (linux-x64 only for now), or use @pdflapas/wasm`, { cause: error });
    }

    const handle = library.symbols.pdflapas_renderer_new();

    if (handle === null) {
        library.close();
        throw new Error("The native renderer could not be created");
    }

    const renderer = new BunNativeRenderer(library, handle, options.onUnsupported, options.onMissingFont);

    for (const [index, font] of options.fonts.entries()) {
        if (library.symbols.pdflapas_add_font(handle, font, font.byteLength) === 0) {
            renderer.close();
            throw new Error(`Font #${index} contains no usable font faces`);
        }
    }

    return renderer;
}
