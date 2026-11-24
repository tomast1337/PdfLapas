// Node/Bun PDFKit renderer ("@pdflapas/pdfkit" under the node/bun conditions): the browser-safe renderer plus
// loading fonts from the system font directories and rendering in worker threads.
import { assertHtmlPresentationDocument, type HtmlPresentationDocument, imagePixelLimits, type MissingFont, type MissingFontPolicy, missingFontPolicy, reportMissingFonts } from "@pdflapas/model";
import { readdir, readFile } from "node:fs/promises";
import { availableParallelism, homedir } from "node:os";
import { extname, join } from "node:path";
import { fileURLToPath } from "node:url";
import { Worker } from "node:worker_threads";
import type { PdfkitRendererPool, PdfkitRendererPoolOptions, SystemFontOptions } from "./browser";
import type { PoolWorkerData, PoolWorkerMessage } from "./nodeWorker";

export {
    createPdfkitRenderer,
    type MissingFont,
    MissingFontError,
    type MissingFontPolicy,
    PdfkitRenderer,
    type PdfkitRendererOptions,
    type PdfkitRendererPool,
    type PdfkitRendererPoolOptions,
    type SystemFontOptions,
    type UnsupportedPolicy,
} from "./browser";

// Caps the default worker count by memory rather than cores: each worker parses the fonts into its own faces.
const MAX_DEFAULT_POOL_WORKERS = 8;

type PoolJob = { readonly document: HtmlPresentationDocument; readonly resolve: (pdf: Uint8Array) => void; readonly reject: (error: unknown) => void };

/** A worker thread and the one document it renders, if any: one document in flight per worker. */
type PoolSlot = { readonly thread: Worker; job: PoolJob | undefined; ready: boolean; crash: Error | undefined };

/**
 * PdfkitRenderers in worker threads (Node/Bun only), so renders use several cores and never block the caller's event
 * loop. Documents wait in one FIFO queue and go to the first free worker. A worker that crashes is replaced and its
 * render rejects. An idle pool does not keep the process alive; `close()` terminates the workers.
 * The same queue as @pdflapas/wasm's WasmWorkerPool, which this package cannot import (tests/packages.test.ts).
 */
export class PdfkitWorkerPool implements PdfkitRendererPool {
    readonly #spawn: () => Worker;
    readonly #onMissingFont: MissingFontPolicy;
    readonly #slots: PoolSlot[] = [];
    readonly #queue: PoolJob[] = [];
    #closed = false;
    #startError: Error | undefined;
    #warnings: readonly string[] = [];
    #missingFonts: readonly MissingFont[] = [];

    /** createPdfkitRendererPool() is the way in: it builds `spawn`, which starts one worker thread on nodeWorker. */
    constructor(spawn: () => Worker, workers: number, onMissingFont?: MissingFontPolicy) {
        this.#spawn = spawn;
        this.#onMissingFont = missingFontPolicy(onMissingFont);
        for (let index = 0; index < workers; index++) {
            this.#start();
        }
    }

    /** Resolves once the workers the pool started with have parsed their fonts; rejects with the first that could not. */
    async started(): Promise<void> {
        const pending: Promise<void>[] = [];

        for (const { thread } of this.#slots) {
            pending.push(
                new Promise((resolve, reject) => {
                    thread.once("message", (message: PoolWorkerMessage) => (message.kind === "failed-to-start" ? reject(message.error) : resolve()));
                    thread.once("exit", code => reject(new Error(`A PdfkitRendererPool worker exited with code ${code} while starting`)));
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
                // As PdfkitRenderer.render() does once the document rendered.
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
                slot.job.reject(new Error(`A PdfkitRendererPool worker exited (code ${code}) while rendering this document; the pool replaced the worker, render the document again`, { cause: slot.crash ?? this.#startError }));
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
            // Structured clone copies the document once; its image data URIs are strings, which cannot be transferred.
            slot.thread.postMessage(job.document);
        }
    }

    #noWorkersError(): Error {
        return new Error(`No PdfkitRendererPool worker could start${this.#startError ? `: ${this.#startError.message}` : ""}`, { cause: this.#startError });
    }

    /** Renders in the first free worker; rejects with the error PdfkitRenderer.render() throws for the same document. */
    async render(document: HtmlPresentationDocument): Promise<Uint8Array> {
        assertHtmlPresentationDocument(document);
        if (this.#closed) {
            throw new Error("This PdfkitRendererPool is closed; create another with createPdfkitRendererPool()");
        }

        if (this.#slots.length === 0) {
            throw this.#noWorkersError();
        }

        // Only the profile's fields: anything else on the object might not clone into the worker.
        const { html, pageSize } = document;

        return new Promise((resolve, reject) => {
            this.#queue.push({ document: { html, pageSize }, resolve, reject });
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
        const closed = new Error("The PdfkitRendererPool was closed before this document rendered");

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
 * Starts `workers` worker threads, each parsing `fonts` into its own PdfkitRenderer, and resolves once all have.
 * Node/Bun only. close() the pool when the service stops; an idle pool does not keep the process alive either way.
 */
export async function createPdfkitRendererPool(options: PdfkitRendererPoolOptions): Promise<PdfkitWorkerPool> {
    if (!options.fonts?.length) {
        throw new Error("createPdfkitRendererPool() requires at least one font: pass { fonts: [Uint8Array, ...] } (await loadSystemFonts())");
    }

    const onUnsupported = options.onUnsupported ?? "error";

    if (onUnsupported !== "error" && onUnsupported !== "warn") {
        throw new Error(`onUnsupported must be "error" or "warn", got ${JSON.stringify(onUnsupported)}`);
    }

    // One worker per core but one, which is left to the event loop that builds the HTML.
    const workers = options.workers ?? Math.max(1, Math.min(availableParallelism() - 1, MAX_DEFAULT_POOL_WORKERS));

    if (!Number.isInteger(workers) || workers < 1) {
        throw new Error(`createPdfkitRendererPool() needs a positive integer number of workers, got ${workers}`);
    }

    // The font bytes once, in shared memory, rather than a copy per worker.
    const fonts: Uint8Array[] = [];

    for (const font of options.fonts) {
        const shared = new Uint8Array(new SharedArrayBuffer(font.byteLength));

        shared.set(font);
        fonts.push(shared);
    }

    // The limits are checked here, so a bad option fails before any worker starts.
    const workerData: PoolWorkerData = { fonts, onUnsupported, imageLimits: imagePixelLimits(options.imageLimits) };
    // The worker's file sits next to this one: nodeWorker.ts beside the sources, nodeWorker.js in dist/.
    const script = new URL(`./nodeWorker${extname(fileURLToPath(import.meta.url))}`, import.meta.url);
    const pool = new PdfkitWorkerPool(() => new Worker(script, { workerData }), workers, options.onMissingFont);

    try {
        await pool.started();
    } catch (error) {
        await pool.close();
        throw error;
    }

    return pool;
}

// The faces the renderer's fallback chain maps Office fonts to (Calibri->Carlito, Cambria->Caladea, Arial->Liberation
// Sans, ... then DejaVu Sans), CJK fallbacks and the LibreOffice target's fallback faces (DejaVu Sans, Serif and Sans
// Mono, Noto Sans and Serif CJK), plus the Office faces themselves: @pdflapas/wasm's set without the
// bitmap emoji font, which has no outlines for PDFKit to embed.
const DEFAULT_FONT_FILES =
    /^(?:(?:Carlito|Caladea|LiberationSans|LiberationSerif|LiberationMono)-.*|DejaVu(?:Sans|Serif|SansMono)(?:-Bold|-Oblique|-BoldOblique|-Italic|-BoldItalic)?|Noto(?:Sans|Serif)CJK-(?:Regular|Bold)|calibri.*|cambria.*|arial.*|times.*|cour.*)\.(?:ttf|otf|ttc)$/i;

async function fontFiles(dir: string, match: RegExp): Promise<string[]> {
    const entries = await readdir(dir, { withFileTypes: true }).catch(() => []);
    const nested = await Promise.all(entries.map(entry => (entry.isDirectory() ? fontFiles(join(dir, entry.name), match) : match.test(entry.name) ? [join(dir, entry.name)] : [])));

    return nested.flat();
}

/** Reads font files from the system font directories, for `createPdfkitRenderer({ fonts })`. Throws when none match. */
export async function loadSystemFonts({ dirs: given, match = DEFAULT_FONT_FILES }: SystemFontOptions = {}): Promise<Uint8Array[]> {
    // The platform's system and user font directories, unless the caller names some.
    let dirs: readonly string[];

    if (given !== undefined) {
        dirs = given;
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
