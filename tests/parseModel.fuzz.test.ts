// Property-based tests for parseModel() (#661, second piece): model JSON built by mutating getModel() either fails with
// InvalidModelError, or is a model whose HTML the WASM renderer draws or reports on. A JavaScript-engine error means a
// value escaped the schema and reached code that did not check it.
import { describe, expect, test } from "bun:test";
import { resolve } from "node:path";
import fc from "fast-check";
import { PuppeteerGen } from "../packages/pptxgenjs/src/html";
import { InvalidModelError, parseModel } from "../packages/pptxgenjs/src/modelSchema";
import { modelToHtml } from "../packages/pptxgenjs/src/render/document";
import type { WasmRenderer } from "../packages/wasm/src/browser";
import { createWasmRenderer, loadSystemFonts } from "../packages/wasm/src/node";
import { FUZZ_SEED, fuzzRuns, fuzzTimeout } from "./fuzzConfig";

const RUNS = fuzzRuns(500);

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/** Messages the JavaScript engine raises for a value it did not expect: an internal failure, not a documented refusal. */
const INTERNAL_FAILURE = [
    /Cannot read propert/,
    /Cannot set propert/,
    /is not a function/,
    /is not iterable/,
    /Cannot convert undefined or null/,
    /Cannot destructure/,
    /Cannot use 'in' operator/,
    /Invalid array length/,
    /Invalid string length/,
    /Maximum call stack size exceeded/,
    /out of memory/i,
    /undefined is not an object/,
    /null is not an object/,
    /Cannot access .* before initialization/,
    /Reduce of empty array/,
];

/** A panic in the Rust renderer aborts the WebAssembly instance: the process (and every later render) is lost. */
const RENDERER_CRASH = /unreachable|panicked|RuntimeError|memory access out of bounds|out of bounds memory/i;

/** Fails when `error` is a JavaScript-engine error (a crash); a documented refusal or a renderer report is accepted. */
function expectNoInternalFailure(error: unknown): void {
    if (!(error instanceof Error)) {
        throw new Error(`an internal call threw a value that is not an Error: ${String(error)}`);
    }

    for (const pattern of INTERNAL_FAILURE) {
        if (pattern.test(error.message)) {
            throw error;
        }
    }
}

/** Fails when the WASM renderer panics; otherwise it either drew the document or reported what it cannot draw. */
function expectRendererReport(error: unknown): void {
    if (!(error instanceof Error)) {
        throw new Error(`the WASM renderer threw a value that is not an Error: ${String(error)}`);
    }

    if (RENDERER_CRASH.test(error.message)) {
        throw error;
    }
}

/** Small valid decks to mutate: text with rich runs and a link, a shape, an image, a table, and a chart. */
function baseModels(): unknown[] {
    const deck = new PuppeteerGen();

    deck.addSlide().addText(
        [
            { text: "Hello ", options: { bold: true, color: "FF0000" } },
            { text: "world", options: { italic: true, hyperlink: { url: "https://example.com", tooltip: "Example" } } },
        ],
        { x: 1, y: 1, w: 4, h: 1, fill: { color: "EEEEEE" }, align: "center", valign: "middle" },
    );
    deck.addSlide().addShape("roundRect", { x: 1, y: 1, w: 2, h: 1, fill: { color: "0088CC" }, rectRadius: 0.2, rotate: 15 });
    deck.addSlide().addImage({ data: PNG, x: 1, y: 1, w: 1, h: 1, transparency: 20 });
    deck.addSlide().addTable(
        [
            [{ text: "A", options: { bold: true } }, { text: "B" }],
            [{ text: "C" }, { text: "D", options: { fill: { color: "DDDDDD" } } }],
        ],
        { x: 1, y: 1, w: 4, colW: [2, 2], border: { type: "solid", pt: 1, color: "000000" }, fontSize: 12 },
    );

    const chart = new PuppeteerGen();

    chart.addSlide().addChart("bar", [{ name: "Series", labels: ["A", "B", "C"], values: [3, 5, 4] }], { x: 1, y: 1, w: 4, h: 3, showLegend: true });

    return [JSON.parse(JSON.stringify(deck.getModel())), JSON.parse(JSON.stringify(chart.getModel()))];
}

// --- Mutating a model's JSON -----------------------------------------------------------------------------------

type PathStep = { readonly key: string } | { readonly index: number };
type Path = readonly PathStep[];

/** Every JSON path in `value`, depth first, to the depth a model nests (a dozen: slides, objects, text, paragraphs, runs). */
function collectPaths(value: unknown, path: Path, out: Path[], depth: number): void {
    out.push(path);

    if (depth >= 12) {
        return;
    }

    if (Array.isArray(value)) {
        for (let index = 0; index < value.length; index++) {
            collectPaths(value[index], [...path, { index }], out, depth + 1);
        }

        return;
    }

    if (typeof value === "object" && value !== null) {
        for (const key of Object.keys(value)) {
            collectPaths((value as Record<string, unknown>)[key], [...path, { key }], out, depth + 1);
        }
    }
}

/** The node `path` names, or undefined when an earlier mutation removed it. */
function nodeAt(root: unknown, path: Path): unknown {
    let node = root;

    for (const step of path) {
        if ("index" in step) {
            node = Array.isArray(node) ? node[step.index] : undefined;
        } else {
            node = typeof node === "object" && node !== null ? (node as Record<string, unknown>)[step.key] : undefined;
        }
    }

    return node;
}

function setAt(root: unknown, path: Path, value: unknown): unknown {
    if (path.length === 0) {
        return value;
    }

    const parent = nodeAt(root, path.slice(0, -1));
    const last = path[path.length - 1];

    if (last === undefined) {
        return root;
    }

    if ("index" in last && Array.isArray(parent)) {
        parent[last.index] = value;
    } else if ("key" in last && typeof parent === "object" && parent !== null && !Array.isArray(parent)) {
        // defineProperty, not assignment: a generated "__proto__" key must become an own JSON key, as JSON.parse makes it.
        Object.defineProperty(parent, last.key, { value, enumerable: true, configurable: true, writable: true });
    }

    return root;
}

function deleteAt(root: unknown, path: Path): unknown {
    if (path.length === 0) {
        return root;
    }

    const parent = nodeAt(root, path.slice(0, -1));
    const last = path[path.length - 1];

    if (last === undefined) {
        return root;
    }

    if ("index" in last && Array.isArray(parent)) {
        parent.splice(last.index, 1);
    } else if ("key" in last && typeof parent === "object" && parent !== null && !Array.isArray(parent)) {
        delete (parent as Record<string, unknown>)[last.key];
    }

    return root;
}

type Mutation =
    | { readonly kind: "set"; readonly path: number; readonly value: unknown }
    | { readonly kind: "delete"; readonly path: number }
    | { readonly kind: "addKey"; readonly path: number; readonly key: string; readonly value: unknown }
    | { readonly kind: "sameType"; readonly path: number; readonly value: unknown }
    | { readonly kind: "deep"; readonly path: number; readonly depth: number }
    | { readonly kind: "none" };

// Values a producer could write into a field: JSON, so no undefined, symbol or function, which no model JSON carries.
const jsonValue = fc.jsonValue({ maxDepth: 2 });
const keyName = fc.oneof(fc.string({ maxLength: 10 }), fc.constantFrom("slides", "pageSize", "kind", "x", "width", "text", "runs", "objects", "__proto__", "constructor", "toString"));
const mutation: fc.Arbitrary<Mutation> = fc.oneof(
    { weight: 5, arbitrary: fc.record({ kind: fc.constant("set" as const), path: fc.nat(), value: jsonValue }) },
    { weight: 2, arbitrary: fc.record({ kind: fc.constant("delete" as const), path: fc.nat() }) },
    { weight: 2, arbitrary: fc.record({ kind: fc.constant("addKey" as const), path: fc.nat(), key: keyName, value: jsonValue }) },
    // A replacement of the same JavaScript type, so more mutations stay inside the schema and reach the renderer.
    { weight: 3, arbitrary: fc.record({ kind: fc.constant("sameType" as const), path: fc.nat(), value: fc.oneof(jsonValue, fc.integer(), fc.constantFrom("", "x", "0".repeat(64), Number.MAX_SAFE_INTEGER, Number.MIN_SAFE_INTEGER, -1, 0, 1, 1e9)) }) },
    // A deeply nested array where a value was, as a producer with a runaway loop would write.
    { weight: 1, arbitrary: fc.record({ kind: fc.constant("deep" as const), path: fc.nat(), depth: fc.integer({ min: 1, max: 40 }) }) },
    { weight: 2, arbitrary: fc.constant({ kind: "none" as const }) },
);
const mutations = fc.array(mutation, { maxLength: 3 });

/** `value` with the same JavaScript type as `target`, from the generated `replacement` when it already has that type. */
function sameTypeValue(target: unknown, replacement: unknown): unknown {
    if (typeof target === "number") {
        return typeof replacement === "number" ? replacement : 1e9;
    }

    if (typeof target === "string") {
        return typeof replacement === "string" ? replacement : JSON.stringify(replacement);
    }

    if (typeof target === "boolean") {
        return typeof replacement === "boolean" ? replacement : true;
    }

    if (Array.isArray(target)) {
        return Array.isArray(replacement) ? replacement : [replacement];
    }

    if (typeof target === "object" && target !== null) {
        return typeof replacement === "object" && replacement !== null && !Array.isArray(replacement) ? replacement : { extra: replacement };
    }

    return replacement;
}

/** The base model with every mutation applied in order; a mutation whose path an earlier one removed does nothing. */
function mutate(base: unknown, operations: readonly Mutation[]): unknown {
    const paths: Path[] = [];

    collectPaths(base, [], paths, 0);

    let model = base;

    for (const operation of operations) {
        if (operation.kind === "none") {
            continue;
        }

        const path = paths[operation.path % paths.length] ?? [];

        if (operation.kind === "set") {
            model = setAt(model, path, operation.value);
        } else if (operation.kind === "delete") {
            model = deleteAt(model, path);
        } else if (operation.kind === "addKey") {
            const parent = nodeAt(model, path);

            if (typeof parent === "object" && parent !== null && !Array.isArray(parent)) {
                setAt(model, [...path, { key: operation.key }], operation.value);
            }
        } else if (operation.kind === "sameType") {
            const target = nodeAt(model, path);

            model = setAt(model, path, sameTypeValue(target, operation.value));
        } else {
            let nested: unknown = [];

            for (let level = 0; level < operation.depth; level++) {
                nested = [nested];
            }

            model = setAt(model, path, nested);
        }
    }

    return model;
}

// --- The property ---------------------------------------------------------------------------------------------

// The binary and fonts the renderer needs; CI's "WASM build" job supplies them (as in tests/wasm.test.ts).
const wasmBuilt = await Bun.file(resolve(import.meta.dir, "../packages/wasm/src/pdf_wasm_bg.wasm")).exists();
const systemFonts = wasmBuilt ? await loadSystemFonts().catch(() => []) : [];
const canRender = wasmBuilt && systemFonts.length > 0;

if (!canRender) {
    const reason = wasmBuilt ? "no system fonts found" : "packages/wasm/src/pdf_wasm_bg.wasm is not built (bun run build:wasm)";

    if (process.env.CI === "true") {
        throw new Error(`parseModel() fuzz cannot run in CI: ${reason}`);
    }

    console.warn(`Skipping the parseModel() renderer property: ${reason}`);
}

const bases = baseModels();
// One renderer for every run: a wasm instance and its fonts cost more than the whole property (AGENTS.md,
// "Performance": set up once, reuse).
const renderer: WasmRenderer | undefined = canRender ? await createWasmRenderer({ fonts: systemFonts, onUnsupported: "warn", onMissingFont: "ignore" }) : undefined;

/**
 * parseModel() on `model`: a model it refuses ends the attempt, a model it accepts must give HTML the renderer draws or
 * reports on. An internal failure (a JavaScript-engine message) is thrown; a renderer report is not.
 */
async function attempt(model: unknown, render: WasmRenderer): Promise<void> {
    let parsed: ReturnType<typeof parseModel>;

    try {
        parsed = parseModel(model);
    } catch (error) {
        if (error instanceof InvalidModelError) {
            return;
        }

        throw error;
    }

    let document: Awaited<ReturnType<typeof modelToHtml>>;

    try {
        document = await modelToHtml(parsed);
    } catch (error) {
        expectNoInternalFailure(error);

        return;
    }

    await render.render(document).catch(expectRendererReport);
}

describe.skipIf(!canRender)("parseModel() fuzz (#661)", () => {
    test(
        "a mutated model is refused with InvalidModelError or its HTML renders, within 2 seconds",
        async () => {
            const render = renderer;

            await fc.assert(
                fc.asyncProperty(fc.nat(), mutations, async (index, operations) => {
                    if (render === undefined) {
                        return;
                    }

                    const base = bases[index % bases.length] ?? {};
                    const started = performance.now();
                    const model = mutate(structuredClone(base), operations);

                    await attempt(model, render);

                    const elapsed = performance.now() - started;

                    // The base decks render in a few milliseconds; the slowest accepted mutation here took well under
                    // 100 ms. 2 s leaves an order of magnitude for a slow CI machine without hiding a case that grows
                    // without bound (a huge page size or length, say).
                    expect(elapsed).toBeLessThan(2000);
                }),
                { seed: FUZZ_SEED, numRuns: RUNS },
            );
        },
        fuzzTimeout(120_000),
    );
});
