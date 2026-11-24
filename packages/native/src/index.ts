// "@pdflapas/native" outside Bun (the default condition): the public types, and a createNativeRenderer() that says
// what to use instead. The renderer itself is in bun.ts, which loads the Rust library with bun:ffi.
import type { MissingFont, MissingFontPolicy, PresentationRenderer } from "@pdflapas/model";

export { type MissingFont, MissingFontError, type MissingFontPolicy } from "@pdflapas/model";

/** The options of `createWasmRenderer` from @pdflapas/wasm that apply to a native library. */
export type NativeRendererOptions = {
    /** TTF/OTF/TTC font files, loaded once into the renderer; at least one. @pdflapas/wasm's `loadSystemFonts()` reads the system's. */
    fonts: readonly Uint8Array[];
    /** As in @pdflapas/wasm: `"error"` (default) throws on what the renderer skips, `"warn"` leaves it in `warnings()`. */
    onUnsupported?: "error" | "warn";
    /** As in @pdflapas/wasm: what to do when a text run asks for a face none of `fonts` provides. */
    onMissingFont?: MissingFontPolicy;
};

/**
 * What createNativeRenderer() returns: the WASM renderer's surface over the same Rust code, plus `close()`, which frees
 * the native memory (fonts included) without waiting for a garbage collection. A render is one synchronous FFI call:
 * it blocks the event loop for its length, as a WasmRenderer render does.
 */
export type NativeRenderer = PresentationRenderer & {
    /** What the last render could not draw. */
    warnings(): string[];
    /** The font faces the last render asked for that none of the renderer's fonts provides. */
    missingFonts(): MissingFont[];
    /** Frees the renderer; later renders throw. */
    close(): void;
    [Symbol.dispose](): void;
};

export type NativeRendererPoolOptions = NativeRendererOptions & {
    /** Render threads, all sharing one copy of `fonts`. Default: one per core but one, at most 8. */
    threads?: number;
};

/**
 * What createNativeRendererPool() returns: the surface of @pdflapas/wasm's WasmRendererPool. Renders run on threads in
 * the library, so `render()` does not block the event loop; only copying the HTML in and the PDF out does.
 */
export type NativeRendererPool = PresentationRenderer & {
    /** What the last render to finish could not draw; with several in flight, read the error of the one you need instead. */
    warnings(): string[];
    /** The font faces the last render to finish asked for that none of the fonts provides. */
    missingFonts(): MissingFont[];
    /** Stops the threads; queued and in-flight renders reject. */
    close(): Promise<void>;
    [Symbol.asyncDispose](): Promise<void>;
};

const NEEDS_BUN = "@pdflapas/native needs Bun's FFI; use @pdflapas/wasm on Node";

export async function createNativeRenderer(_options: NativeRendererOptions): Promise<NativeRenderer> {
    throw new Error(NEEDS_BUN);
}

export async function createNativeRendererPool(_options: NativeRendererPoolOptions): Promise<NativeRendererPool> {
    throw new Error(NEEDS_BUN);
}
