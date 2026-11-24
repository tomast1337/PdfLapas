// "@pdflapas/native" under Node (the node condition): the Bun entry's names, so a codebase shared between Bun and Node
// can import it, with factories that throw when called and say what to use on Node.
//
// shortcut: no Node renderer until node:ffi leaves Experimental (https://nodejs.org/api/ffi.html, Node 26.1+, Stability
// 1); port then (#612 phase 4). node:ffi has dlopen(path, { fn: { arguments: [...], return } }), suffix and
// toArrayBuffer(ptr, len, copy), so the single renderer maps almost 1:1 from bun.ts. The pool does not: a
// registerCallback callback must run on the thread that created it, so the threadsafe `wake` doorbell bun.ts hands
// pdflapas_pool_new cannot be used; pass null and drain pdflapas_pool_next on a setImmediate/timer loop while renders are
// in flight. node:ffi is also disabled under the Permission Model unless the process runs with --allow-ffi.
import { type NativeRenderer, type NativeRendererOptions, type NativeRendererPool, type NativeRendererPoolOptions, UnsupportedRuntimeError } from "./index";

export type { MissingFont, MissingFontPolicy, NativeRenderer, NativeRendererOptions, NativeRendererPool, NativeRendererPoolOptions } from "./index";
export { MissingFontError, UnsupportedRuntimeError } from "./index";

const NEEDS_BUN_ON_NODE = "@pdflapas/native needs Bun (bun:ffi); on Node use @pdflapas/wasm (createWasmRenderer / createWasmRendererPool). Node support waits on node:ffi (#612 phase 4).";

export async function createNativeRenderer(_options: NativeRendererOptions): Promise<NativeRenderer> {
    throw new UnsupportedRuntimeError("node", NEEDS_BUN_ON_NODE);
}

export async function createNativeRendererPool(_options: NativeRendererPoolOptions): Promise<NativeRendererPool> {
    throw new UnsupportedRuntimeError("node", NEEDS_BUN_ON_NODE);
}
