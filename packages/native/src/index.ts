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

export async function createNativeRenderer(_options: NativeRendererOptions): Promise<NativeRenderer> {
    throw new Error("@pdflapas/native needs Bun's FFI; use @pdflapas/wasm on Node");
}
