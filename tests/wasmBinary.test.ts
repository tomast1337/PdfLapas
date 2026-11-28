import { test } from "bun:test";
import { resolve } from "node:path";
import { checkWasmBinary } from "../scripts/publish";

// A binary built from older Rust sources fails unrelated WASM tests in ways that take long to trace;
// this names the cause. Without a binary the WASM tests skip, so this does too.
const wasmBuilt = await Bun.file(resolve(import.meta.dir, "../packages/wasm/src/pdf_wasm_bg.wasm")).exists();

test.skipIf(!wasmBuilt)("the local WASM binary is the one the committed sources build", async () => {
    await checkWasmBinary();
});
