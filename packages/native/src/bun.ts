// "@pdflapas/native" under Bun: the Rust renderer (crates/pdf-ffi, pdf-core) as a native library loaded with bun:ffi.
// Same surface, options and errors as @pdflapas/wasm's WasmRenderer and WasmRendererPool.
import { assertHtmlPresentationDocument, type HtmlPresentationDocument, type MissingFont, type MissingFontPolicy, missingFontPolicy, reportMissingFonts } from "@pdflapas/model";
import { dlopen, FFIType, JSCallback, type Library, type Pointer, suffix, toArrayBuffer } from "bun:ffi";
import { availableParallelism } from "node:os";
import { fileURLToPath } from "node:url";
import type { NativeRenderer, NativeRendererOptions, NativeRendererPool, NativeRendererPoolOptions } from "./index";

export type { MissingFont, MissingFontPolicy, NativeRenderer, NativeRendererOptions, NativeRendererPool, NativeRendererPoolOptions } from "./index";
export { MissingFontError, UnsupportedRuntimeError } from "./index";

/** The C ABI of crates/pdf-ffi (lib.rs, pool.rs). `usize` is `u64`: the prototype builds 64-bit Linux only. */
const SYMBOLS = {
    pdflapas_renderer_new: { args: [], returns: FFIType.ptr },
    pdflapas_add_font: { args: [FFIType.ptr, FFIType.ptr, FFIType.u64], returns: FFIType.u32 },
    pdflapas_render: { args: [FFIType.ptr, FFIType.ptr, FFIType.u64, FFIType.ptr], returns: FFIType.i32 },
    pdflapas_report: { args: [FFIType.ptr, FFIType.ptr], returns: FFIType.i32 },
    pdflapas_last_error: { args: [FFIType.ptr, FFIType.ptr], returns: FFIType.i32 },
    pdflapas_buffer_free: { args: [FFIType.ptr, FFIType.u64], returns: FFIType.void },
    pdflapas_renderer_free: { args: [FFIType.ptr], returns: FFIType.void },
    pdflapas_pool_new: { args: [FFIType.u32, FFIType.function], returns: FFIType.ptr },
    pdflapas_pool_add_font: { args: [FFIType.ptr, FFIType.ptr, FFIType.u64], returns: FFIType.u32 },
    pdflapas_pool_start: { args: [FFIType.ptr], returns: FFIType.i32 },
    pdflapas_pool_submit: { args: [FFIType.ptr, FFIType.u64, FFIType.ptr, FFIType.u64], returns: FFIType.i32 },
    pdflapas_pool_next: { args: [FFIType.ptr, FFIType.ptr], returns: FFIType.i32 },
    pdflapas_pool_free: { args: [FFIType.ptr], returns: FFIType.void },
} as const;

type NativeLibrary = Library<typeof SYMBOLS>;

/** `PDFLAPAS_OK` in crates/pdf-ffi. */
const OK = 0;
/** `pdflapas_pool_next` wrote a completion. */
const COMPLETED = 1;
/** u64 fields of a `PdflapasCompletion`: id, status, then (ptr, len) of pdf, report and error. */
const COMPLETION_FIELDS = 8;
// Caps the default thread count. Threads share the font bytes, so the cap is not about memory as the WASM pool's is;
// past 8 the bench:load mix is bound by the event loop building the HTML, not by rendering.
const MAX_DEFAULT_POOL_THREADS = 8;
// While renders are in flight, a timer keeps the process alive (a threadsafe JSCallback does not) and drains the queue
// as a safety net should a wake-up ever be lost.
const IN_FLIGHT_DRAIN_MS = 1000;
const ENCODER = new TextEncoder();
const DECODER = new TextDecoder();

/** `pdflapas_report`'s JSON: `warnings()` and `missingFonts()` of crates/pdf-wasm. */
type Report = { readonly warnings: string[]; readonly missingFonts: readonly { readonly face: string; readonly used?: string; readonly objects: string[] }[] };

/** Loads the platform's library next to this module. */
function openLibrary(): NativeLibrary {
    // shortcut: built for linux-x64 only (bun run build:native); per-platform packages are phase 2 (#612).
    const path = fileURLToPath(new URL(`./pdflapas_native.${process.platform}-${process.arch}.${suffix}`, import.meta.url));

    try {
        return dlopen(path, SYMBOLS);
    } catch (error) {
        throw new Error(`No @pdflapas/native library for ${process.platform}-${process.arch} at ${path}: build it with \`bun run build:native\` (linux-x64 only for now), or use @pdflapas/wasm`, { cause: error });
    }
}

/** Copies a buffer the library handed over into JS memory and frees it in the library. */
function take(library: NativeLibrary, address: bigint, length: bigint): Uint8Array {
    // A Rust pointer; Number keeps every user-space x86-64 address exactly.
    const pointer = Number(address) as Pointer;
    const bytes = length === 0n ? new Uint8Array() : new Uint8Array(toArrayBuffer(pointer, 0, Number(length))).slice();

    library.symbols.pdflapas_buffer_free(pointer, length);

    return bytes;
}

/** The warnings and missing fonts of one render, from `pdflapas_report`'s JSON. */
function readReport(json: Uint8Array): { warnings: string[]; missingFonts: MissingFont[] } {
    // The library's own output, built by crates/pdf-ffi's report_json.
    const report = JSON.parse(DECODER.decode(json)) as Report;
    const missingFonts: MissingFont[] = [];

    for (const { face, used, objects } of report.missingFonts) {
        missingFonts.push({ face, objects, ...(used ? { usedFace: used } : {}) });
    }

    return { warnings: report.warnings, missingFonts };
}

/** What `onUnsupported: "error"` throws, as WasmRenderer does. */
function unsupportedError(warnings: readonly string[]): Error {
    return new Error(`The native renderer cannot draw everything in this presentation (pass { onUnsupported: "warn" } to render anyway):\n- ${warnings.join("\n- ")}`);
}

/** One renderer: each render is one synchronous FFI call that blocks the event loop, as a WasmRenderer render does. */
class BunNativeRenderer implements NativeRenderer {
    readonly #library: NativeLibrary;
    readonly #onUnsupported: "error" | "warn";
    readonly #onMissingFont: MissingFontPolicy;
    /** The `PdflapasBuffer` (`ptr`, `len`) every call writes its output to, reused. */
    readonly #out = new BigUint64Array(2);
    #handle: Pointer | null;
    #warnings: string[] = [];
    #missingFonts: MissingFont[] = [];

    constructor(library: NativeLibrary, handle: Pointer, onUnsupported: "error" | "warn" = "error", onMissingFont?: MissingFontPolicy) {
        this.#library = library;
        this.#handle = handle;
        this.#onUnsupported = onUnsupported;
        this.#onMissingFont = missingFontPolicy(onMissingFont);
    }

    /** Parses the profile HTML in Rust; out-of-profile or malformed HTML throws, naming slide, object and property. Images must be data URIs. */
    async render(document: HtmlPresentationDocument): Promise<Uint8Array> {
        assertHtmlPresentationDocument(document);
        if (this.#handle === null) {
            throw new Error("This native renderer is closed; create another with createNativeRenderer()");
        }

        const handle = this.#handle;
        const html = ENCODER.encode(document.html);
        const { symbols } = this.#library;
        const out = this.#out;

        if (symbols.pdflapas_render(handle, html, html.byteLength, out) !== OK) {
            symbols.pdflapas_last_error(handle, out);
            throw new Error(DECODER.decode(take(this.#library, out[0] ?? 0n, out[1] ?? 0n)));
        }

        const pdf = take(this.#library, out[0] ?? 0n, out[1] ?? 0n);

        if (symbols.pdflapas_report(handle, out) !== OK) {
            symbols.pdflapas_last_error(handle, out);
            throw new Error(DECODER.decode(take(this.#library, out[0] ?? 0n, out[1] ?? 0n)));
        }

        ({ warnings: this.#warnings, missingFonts: this.#missingFonts } = readReport(take(this.#library, out[0] ?? 0n, out[1] ?? 0n)));
        if (this.#warnings.length && this.#onUnsupported !== "warn") {
            throw unsupportedError(this.#warnings);
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
}

/** Loads the platform's library next to this module and the fonts into one renderer. Bun only. */
export async function createNativeRenderer(options: NativeRendererOptions): Promise<NativeRenderer> {
    if (!options.fonts?.length) {
        throw new Error("createNativeRenderer() requires at least one font: pass { fonts: [Uint8Array, ...] } (@pdflapas/wasm's loadSystemFonts() reads the system's)");
    }

    const library = openLibrary();
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

type PoolJob = { readonly resolve: (pdf: Uint8Array) => void; readonly reject: (error: unknown) => void };

/**
 * Render threads in the library behind one handle (crates/pdf-ffi/src/pool.rs). `render()` copies the HTML into the
 * library's queue and returns; a thread renders it; the library rings `wake` (a threadsafe JSCallback) and `drain()`,
 * on the event loop, takes every finished job and settles its promise.
 */
class BunNativePool implements NativeRendererPool {
    readonly #library: NativeLibrary;
    readonly #onUnsupported: "error" | "warn";
    readonly #onMissingFont: MissingFontPolicy;
    /** The `PdflapasCompletion` `pdflapas_pool_next` writes, reused. */
    readonly #completion = new BigUint64Array(COMPLETION_FIELDS);
    readonly #pending = new Map<number, PoolJob>();
    #handle: Pointer | null;
    #nextId = 0;
    #inFlight: ReturnType<typeof setInterval> | undefined;
    #warnings: string[] = [];
    #missingFonts: MissingFont[] = [];

    constructor(library: NativeLibrary, handle: Pointer, onUnsupported: "error" | "warn" = "error", onMissingFont?: MissingFontPolicy) {
        this.#library = library;
        this.#handle = handle;
        this.#onUnsupported = onUnsupported;
        this.#onMissingFont = missingFontPolicy(onMissingFont);
    }

    /** Queues the document for the first free thread; rejects with the error a NativeRenderer throws for it. */
    async render(document: HtmlPresentationDocument): Promise<Uint8Array> {
        assertHtmlPresentationDocument(document);
        if (this.#handle === null) {
            throw new Error("This native renderer pool is closed; create another with createNativeRendererPool()");
        }

        const html = ENCODER.encode(document.html);
        const id = this.#nextId++;

        if (this.#library.symbols.pdflapas_pool_submit(this.#handle, id, html, html.byteLength) !== OK) {
            throw new Error("The native renderer pool could not queue this document");
        }

        this.#inFlight ??= setInterval(() => this.drain(), IN_FLIGHT_DRAIN_MS);

        return new Promise((resolve, reject) => {
            this.#pending.set(id, { resolve, reject });
        });
    }

    /**
     * Settles the oldest finished job and, if it got one, comes back for the next in a later macrotask, as a worker's
     * messages arrive one per task: settling a whole batch at once would run every caller's continuation before any timer
     * or I/O. Called by the library's wake-up and the in-flight timer; never throws.
     */
    drain(): void {
        if (this.#handle === null || this.#library.symbols.pdflapas_pool_next(this.#handle, this.#completion) !== COMPLETED) {
            if (this.#pending.size === 0) {
                clearInterval(this.#inFlight);
                this.#inFlight = undefined;
            }

            return;
        }

        setImmediate(() => this.drain());
        const [id = 0n, status = 0n, pdfAt = 0n, pdfLength = 0n, reportAt = 0n, reportLength = 0n, errorAt = 0n, errorLength = 0n] = this.#completion;
        const pdf = take(this.#library, pdfAt, pdfLength);
        const report = readReport(take(this.#library, reportAt, reportLength));
        const message = DECODER.decode(take(this.#library, errorAt, errorLength));
        const job = this.#pending.get(Number(id));

        this.#pending.delete(Number(id));
        ({ warnings: this.#warnings, missingFonts: this.#missingFonts } = report);
        if (job === undefined) {
            return;
        }

        if (status !== BigInt(OK)) {
            job.reject(new Error(message));
        } else if (report.warnings.length && this.#onUnsupported !== "warn") {
            job.reject(unsupportedError(report.warnings));
        } else {
            try {
                reportMissingFonts(report.missingFonts, this.#onMissingFont);
                job.resolve(pdf);
            } catch (error) {
                job.reject(error);
            }
        }
    }

    /** What the last render to finish could not draw; with several in flight, read the error of the one you need instead. */
    warnings(): string[] {
        return [...this.#warnings];
    }

    /** The font faces the last render to finish asked for that none of the fonts provides. */
    missingFonts(): MissingFont[] {
        return [...this.#missingFonts];
    }

    /** Stops the threads (renders in flight finish first, queued ones are dropped); pending renders reject. */
    async close(): Promise<void> {
        if (this.#handle === null) {
            return;
        }

        // Joins the threads: no wake-up comes after this.
        this.#library.symbols.pdflapas_pool_free(this.#handle);
        this.#handle = null;
        this.#library.close();
        clearInterval(this.#inFlight);
        this.#inFlight = undefined;
        const closed = new Error("The native renderer pool was closed before this document rendered");

        for (const job of this.#pending.values()) {
            job.reject(closed);
        }

        this.#pending.clear();
    }

    async [Symbol.asyncDispose](): Promise<void> {
        await this.close();
    }
}

/**
 * Starts `threads` render threads in the library behind one handle, the font files read once and shared by every
 * thread. `render()` returns at once and resolves on the event loop when a thread has drawn the PDF. Bun only.
 */
export async function createNativeRendererPool(options: NativeRendererPoolOptions): Promise<NativeRendererPool> {
    if (!options.fonts?.length) {
        throw new Error("createNativeRendererPool() requires at least one font: pass { fonts: [Uint8Array, ...] } (@pdflapas/wasm's loadSystemFonts() reads the system's)");
    }

    // One thread per core but one, which is left to the event loop that builds the HTML.
    const threads = options.threads ?? Math.max(1, Math.min(availableParallelism() - 1, MAX_DEFAULT_POOL_THREADS));

    if (!Number.isInteger(threads) || threads < 1) {
        throw new Error(`createNativeRendererPool() needs a positive integer number of threads, got ${threads}`);
    }

    const library = openLibrary();
    let pool: BunNativePool | undefined;
    // A doorbell only: it carries no data, so a late or coalesced call is harmless. It lives as long as the process
    // (shortcut: closing it races a wake-up Bun may already have queued; one callback per pool, pools are long-lived).
    const wake = new JSCallback(() => pool?.drain(), { args: [], returns: FFIType.void, threadsafe: true });
    const handle = library.symbols.pdflapas_pool_new(threads, wake);

    if (handle === null) {
        library.close();
        throw new Error("The native renderer pool could not be created");
    }

    pool = new BunNativePool(library, handle, options.onUnsupported, options.onMissingFont);
    for (const [index, font] of options.fonts.entries()) {
        if (library.symbols.pdflapas_pool_add_font(handle, font, font.byteLength) === 0) {
            await pool.close();
            throw new Error(`Font #${index} contains no usable font faces`);
        }
    }

    if (library.symbols.pdflapas_pool_start(handle) !== OK) {
        await pool.close();
        throw new Error(`The native renderer pool could not start ${threads} threads`);
    }

    return pool;
}
