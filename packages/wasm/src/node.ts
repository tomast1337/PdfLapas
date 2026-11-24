// Node/Bun WASM text renderer ("@pdflapas/wasm" under the node/bun conditions):
// reads the .wasm next to this module, can load fonts from the system font directories, and renders in worker threads.
import { assertHtmlPresentationDocument, type HtmlPresentationDocument, type MissingFont, type ModelHtmlRequest, type MissingFontPolicy, missingFontPolicy, reportMissingFonts } from "@pdflapas/model";
import { readdir, readFile } from "node:fs/promises";
import { availableParallelism, homedir } from "node:os";
import { extname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { Worker } from "node:worker_threads";
import { createWasmRenderer as createBrowserWasmRenderer, type SystemFontOptions, type WasmRenderer, type WasmRendererOptions, type WasmRendererPool, type WasmRendererPoolOptions } from "./browser";
import type { PoolRequest, PoolWorkerData, PoolWorkerMessage } from "./nodeWorker";

export type { MissingFontBinding, SystemFontOptions, TextPdfBinding, WasmRendererOptions, WasmRendererPool, WasmRendererPoolOptions, WasmSource } from "./browser";
export { type MissingFont, MissingFontError, type MissingFontPolicy } from "./browser";
export { WasmRenderer } from "./browser";

/** The `pdf_wasm_bg.wasm` next to this module. */
async function readWasmBinary(): Promise<Uint8Array> {
    const path = fileURLToPath(new URL("./pdf_wasm_bg.wasm", import.meta.url));

    return readFile(path).catch((error: unknown) => {
        throw new Error(`WASM renderer binary not found at ${path}; build it with \`bun run build:wasm\` or pass { wasm }`, { cause: error });
    });
}

export async function createWasmRenderer(options: WasmRendererOptions): Promise<WasmRenderer> {
    // Without fonts, let the shared loader report that before touching the filesystem.
    if (options.wasm !== undefined || !options.fonts?.length) {
        return createBrowserWasmRenderer(options);
    }

    return createBrowserWasmRenderer({ ...options, wasm: await readWasmBinary() });
}

// Caps the default worker count by memory rather than cores: each worker holds its own copy of the fonts.
const MAX_DEFAULT_POOL_WORKERS = 8;

type PoolJob = { readonly request: PoolRequest; readonly resolve: (pdf: Uint8Array) => void; readonly reject: (error: unknown) => void };

/** A worker thread and the one document it renders, if any: one document in flight per worker. */
type PoolSlot = { readonly thread: Worker; job: PoolJob | undefined; ready: boolean; crash: Error | undefined };

/**
 * WasmRenderers in worker threads (Node/Bun only), so renders use several cores and never block the caller's event
 * loop. Documents wait in one FIFO queue and go to the first free worker. A worker that crashes is replaced and its
 * render rejects. An idle pool does not keep the process alive; `close()` terminates the workers.
 */
export class WasmWorkerPool implements WasmRendererPool {
    readonly #spawn: () => Worker;
    readonly #onMissingFont: MissingFontPolicy;
    readonly #slots: PoolSlot[] = [];
    readonly #queue: PoolJob[] = [];
    /** Set only with `buildHtmlInWorkers`: absent, toPdf() builds the HTML on its own thread and no worker loads the builder (#634). */
    readonly renderModelHtml?: (request: ModelHtmlRequest) => Promise<Uint8Array>;
    #closed = false;
    #startError: Error | undefined;
    #warnings: readonly string[] = [];
    #missingFonts: readonly MissingFont[] = [];

    /** createWasmRendererPool() is the way in: it builds `spawn`, which starts one worker thread on nodeWorker. */
    constructor(spawn: () => Worker, workers: number, onMissingFont?: MissingFontPolicy, buildHtmlInWorkers = false) {
        this.#spawn = spawn;
        this.#onMissingFont = missingFontPolicy(onMissingFont);
        if (buildHtmlInWorkers) {
            // Builds the HTML from the model in the first free worker, with `modelHtml` from `request.htmlModule`, then renders it there; same bytes as render().
            this.renderModelHtml = async ({ model, images, htmlModule }) => this.#enqueue({ kind: "model", model, images, htmlModule });
        }

        for (let index = 0; index < workers; index++) {
            this.#start();
        }
    }

    /** Resolves once the workers the pool started with have loaded their fonts; rejects with the first that could not. */
    async started(): Promise<void> {
        const pending: Promise<void>[] = [];

        for (const { thread } of this.#slots) {
            pending.push(
                new Promise((resolve, reject) => {
                    thread.once("message", (message: PoolWorkerMessage) => (message.kind === "failed-to-start" ? reject(message.error) : resolve()));
                    thread.once("exit", code => reject(new Error(`A WasmRendererPool worker exited with code ${code} while starting`)));
                }),
            );
        }

        await Promise.all(pending);
    }

    #start(): void {
        const slot: PoolSlot = { thread: this.#spawn(), job: undefined, ready: false, crash: undefined };

        this.#slots.push(slot);
        slot.thread.on("message", (message: PoolWorkerMessage) => {
            if (message.kind === "ready") {
                slot.ready = true;
                if (slot.job === undefined) {
                    slot.thread.unref();
                }

                return;
            }

            if (message.kind === "failed-to-start") {
                this.#startError = message.error;

                return;
            }

            const job = slot.job;

            slot.job = undefined;
            slot.thread.unref();
            this.#warnings = message.warnings;
            this.#missingFonts = message.missingFonts;
            if (message.kind === "error") {
                job?.reject(message.error);
            } else {
                // As WasmRenderer.render() does once the document rendered.
                try {
                    reportMissingFonts(message.missingFonts, this.#onMissingFont);
                    job?.resolve(message.pdf);
                } catch (error) {
                    job?.reject(error);
                }
            }

            this.#pump();
        });
        // An uncaught exception in the worker; "exit" follows.
        slot.thread.on("error", error => {
            slot.crash = error;
        });
        slot.thread.on("exit", code => {
            this.#slots.splice(this.#slots.indexOf(slot), 1);
            if (slot.job !== undefined) {
                slot.job.reject(new Error(`A WasmRendererPool worker exited (code ${code}) while rendering this document; the pool replaced the worker, render the document again`, { cause: slot.crash ?? this.#startError }));
                slot.job = undefined;
            }

            if (this.#closed) {
                return;
            }

            // A worker that never started would fail again the same way: replace only one that had started.
            if (slot.ready) {
                this.#start();
            } else if (this.#slots.length === 0) {
                for (const job of this.#queue.splice(0)) {
                    job.reject(this.#noWorkersError());
                }
            }

            this.#pump();
        });
    }

    /** Hands queued documents to free workers, oldest first. */
    #pump(): void {
        for (const slot of this.#slots) {
            if (slot.job !== undefined) {
                continue;
            }

            const job = this.#queue.shift();

            if (job === undefined) {
                return;
            }

            slot.job = job;
            // A render in flight keeps the process alive until it settles.
            slot.thread.ref();
            // Structured clone copies the request once; the image data URIs are strings, which cannot be transferred.
            slot.thread.postMessage(job.request);
        }
    }

    #noWorkersError(): Error {
        return new Error(`No WasmRendererPool worker could start${this.#startError ? `: ${this.#startError.message}` : ""}`, { cause: this.#startError });
    }

    /** Renders in the first free worker; rejects with the error WasmRenderer.render() throws for the same document. */
    async render(document: HtmlPresentationDocument): Promise<Uint8Array> {
        assertHtmlPresentationDocument(document);
        // Only the profile's fields: anything else on the object might not clone into the worker.
        const { html, pageSize } = document;

        return this.#enqueue({ kind: "document", document: { html, pageSize } });
    }

    #enqueue(request: PoolRequest): Promise<Uint8Array> {
        if (this.#closed) {
            return Promise.reject(new Error("This WasmRendererPool is closed; create another with createWasmRendererPool()"));
        }

        if (this.#slots.length === 0) {
            return Promise.reject(this.#noWorkersError());
        }

        return new Promise((resolve, reject) => {
            this.#queue.push({ request, resolve, reject });
            this.#pump();
        });
    }

    /** What the last render to finish could not draw; with several in flight, read the error of the one you need instead. */
    warnings(): string[] {
        return [...this.#warnings];
    }

    /** The font faces the last render to finish asked for that none of the fonts provides. */
    missingFonts(): MissingFont[] {
        return [...this.#missingFonts];
    }

    /** Terminates the workers; queued and in-flight renders reject. */
    async close(): Promise<void> {
        this.#closed = true;
        const closed = new Error("The WasmRendererPool was closed before this document rendered");

        for (const job of this.#queue.splice(0)) {
            job.reject(closed);
        }

        const exits: Promise<number>[] = [];

        for (const slot of this.#slots) {
            slot.job?.reject(closed);
            slot.job = undefined;
            exits.push(slot.thread.terminate());
        }

        await Promise.all(exits);
    }

    async [Symbol.asyncDispose](): Promise<void> {
        await this.close();
    }
}

/**
 * Starts `workers` worker threads, each loading `fonts` into its own WasmRenderer, and resolves once all have. Node/Bun
 * only. close() the pool when the service stops; an idle pool does not keep the process alive either way.
 */
export async function createWasmRendererPool(options: WasmRendererPoolOptions): Promise<WasmWorkerPool> {
    if (!options.fonts?.length) {
        throw new Error("createWasmRendererPool() requires at least one font: pass { fonts: [Uint8Array, ...] } (await loadSystemFonts())");
    }

    if (options.buildHtmlInWorkers !== undefined && typeof options.buildHtmlInWorkers !== "boolean") {
        throw new Error(`createWasmRendererPool() needs buildHtmlInWorkers to be true or false, got ${String(options.buildHtmlInWorkers)}`);
    }

    // One worker per core but one, which is left to the event loop that builds the HTML.
    const workers = options.workers ?? Math.max(1, Math.min(availableParallelism() - 1, MAX_DEFAULT_POOL_WORKERS));

    if (!Number.isInteger(workers) || workers < 1) {
        throw new Error(`createWasmRendererPool() needs a positive integer number of workers, got ${workers}`);
    }

    // Compiled once here; each worker only instantiates it.
    const wasm = await WebAssembly.compile(await readWasmBinary());
    // The font bytes once, in shared memory, rather than a copy per worker on top of the copy each renderer keeps.
    const fonts: Uint8Array[] = [];

    for (const font of options.fonts) {
        const shared = new Uint8Array(new SharedArrayBuffer(font.byteLength));

        shared.set(font);
        fonts.push(shared);
    }

    const workerData: PoolWorkerData = { fonts, wasm, onUnsupported: options.onUnsupported ?? "error" };
    // The worker's file sits next to this one: nodeWorker.ts beside the sources, nodeWorker.js in dist/.
    const script = new URL(`./nodeWorker${extname(fileURLToPath(import.meta.url))}`, import.meta.url);
    const pool = new WasmWorkerPool(() => new Worker(script, { workerData }), workers, options.onMissingFont, options.buildHtmlInWorkers);

    try {
        await pool.started();
    } catch (error) {
        await pool.close();
        throw error;
    }

    return pool;
}

// The metric-compatible faces the renderer's fallback chain maps Office fonts to (Calibri->Carlito,
// Cambria->Caladea, Arial->Liberation Sans, ... then DejaVu Sans, Liberation Sans), the LibreOffice target's fallback faces (Noto Color
// Emoji, DejaVu Sans, Serif and Sans Mono, Noto Sans, Serif and Sans Mono CJK), plus the Office faces themselves
// and, where installed, the PowerPoint target's fallback faces (Cambria Math is in cambria.ttc; Segoe UI Symbol and Emoji, Yu Gothic
// regular and bold, MS Gothic, Microsoft JhengHei and YaHei, Malgun Gothic, SimSun, PMingLiU in mingliu.ttc).
const DEFAULT_FONT_FILES =
    /^(?:(?:Carlito|Caladea|LiberationSans|LiberationSerif|LiberationMono)-.*|DejaVu(?:Sans|Serif|SansMono)(?:-Bold|-Oblique|-BoldOblique|-Italic|-BoldItalic)?|NotoColorEmoji|Noto(?:Sans|Serif)CJK-(?:Regular|Bold)|calibri.*|cambria.*|arial.*|times.*|cour.*|seguisym|seguiemj|YuGoth[RB]|msgothic|msjh(?:bd)?|msyh(?:bd)?|malgun(?:bd)?|simsun|mingliu)\.(?:ttf|otf|ttc)$/i;

async function fontFiles(dir: string, match: RegExp): Promise<string[]> {
    const entries = await readdir(dir, { withFileTypes: true }).catch(() => []);
    const nested = await Promise.all(entries.map(entry => (entry.isDirectory() ? fontFiles(join(dir, entry.name), match) : match.test(entry.name) ? [join(dir, entry.name)] : [])));

    return nested.flat();
}

/** Reads font files from the system font directories, for `createWasmRenderer({ fonts })`. Throws when none match. */
export async function loadSystemFonts({ dirs: requestedDirs, match = DEFAULT_FONT_FILES }: SystemFontOptions = {}): Promise<Uint8Array[]> {
    // The platform's system and user font directories, unless the caller named its own.
    let dirs: readonly string[];

    if (requestedDirs !== undefined) {
        dirs = requestedDirs;
    } else if (process.platform === "win32") {
        dirs = [join(process.env.WINDIR ?? "C:\\Windows", "Fonts"), join(process.env.LOCALAPPDATA ?? "", "Microsoft/Windows/Fonts")];
    } else if (process.platform === "darwin") {
        dirs = ["/System/Library/Fonts", "/Library/Fonts", join(homedir(), "Library/Fonts")];
    } else {
        dirs = ["/usr/share/fonts", "/usr/local/share/fonts", join(homedir(), ".local/share/fonts"), join(homedir(), ".fonts")];
    }

    const files = (await Promise.all(dirs.map(dir => fontFiles(dir, match)))).flat().sort();

    if (!files.length) {
        throw new Error(`No font files matching ${match} in ${dirs.join(", ")}; install Carlito, Caladea, Liberation and DejaVu fonts or pass { dirs, match }`);
    }

    return Promise.all(files.map(async file => new Uint8Array(await readFile(file))));
}
