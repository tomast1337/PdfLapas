// Sustained concurrent load (#406): N requests in flight for a fixed time, each building a deck from the mix and
// rendering it through toPdf(), as a service does. Puppeteer uses the Node entry's default renderer (one shared
// Chromium through withBrowser); WASM and PDFKit use one long-lived renderer for every request, wasm-pool one
// long-lived createWasmRendererPool(). Prints, every sample
// interval, per-render latency (p50, p95, p99), throughput, the process's RSS and heap, and the count and RSS of its
// child processes (Chromium); after each level, the same once the load has stopped and the idle browser has closed.
import { readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import type { PresentationRenderer } from "../packages/model/src/index";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { PAGE_SIZES } from "../packages/pptxgenjs/src/pageLayouts";
import { DECKS } from "../tests/visual/wasm/decks";
import { benchLoadConfig } from "./benchConfig";

/** Nearest-rank percentile of ascending `sorted`, `fraction` in [0, 1]. */
export function percentile(sorted: readonly [number, ...number[]], fraction: number): number {
    const index = Math.min(sorted.length - 1, Math.max(0, Math.ceil(fraction * sorted.length) - 1));

    return sorted[index] ?? sorted[0];
}

/** Every process below `root` in a pid -> parent pid table, at any depth. */
export function descendants(parents: ReadonlyMap<number, number>, root: number): number[] {
    const children = new Map<number, number[]>();

    for (const [pid, parent] of parents) {
        children.set(parent, [...(children.get(parent) ?? []), pid]);
    }

    const found: number[] = [];
    const pending = [root];

    for (let pid = pending.pop(); pid !== undefined; pid = pending.pop()) {
        for (const child of children.get(pid) ?? []) {
            found.push(child);
            pending.push(child);
        }
    }

    return found;
}

const PPID_LINE = /^PPid:\s+(\d+)$/m;
const RSS_LINE = /^VmRSS:\s+(\d+) kB$/m;
const PID_NAME = /^\d+$/;
const KB_PER_MB = 1024;
const BYTES_PER_MB = 1024 * 1024;
const MS_PER_S = 1000;
// Longer than IDLE_CLOSE_MS in packages/puppeteer/src/browser.ts, so the shared browser has closed when the after-load sample runs.
const SETTLE_MS = 3000;
// A PNG cut after its IHDR chunk: its header passes checkImagePixels, so the render fails inside the renderer (in
// Chromium's browser context for Puppeteer) and exercises the error path under load.
const BROKEN_PNG_BYTES = 40;
const BROKEN_RENDERED = "the broken image rendered without an error";

/** This process's RSS and heap, and the count and RSS of its descendants (Chromium and its helpers) from /proc. */
function memoryLine(): string {
    const { rss, heapUsed } = process.memoryUsage();
    const status = new Map<number, string>();
    let browser = "children unavailable (no /proc)";

    try {
        for (const entry of readdirSync("/proc")) {
            if (PID_NAME.test(entry)) {
                try {
                    status.set(Number(entry), readFileSync(join("/proc", entry, "status"), "utf8"));
                } catch {
                    // The process exited between the listing and the read.
                }
            }
        }

        const parents = new Map<number, number>();

        for (const [pid, text] of status) {
            const ppid = PPID_LINE.exec(text)?.[1];

            if (ppid !== undefined) {
                parents.set(pid, Number(ppid));
            }
        }

        const pids = descendants(parents, process.pid);
        let rssKb = 0;

        for (const pid of pids) {
            rssKb += Number(RSS_LINE.exec(status.get(pid) ?? "")?.[1] ?? 0);
        }

        browser = `children ${pids.length} procs ${(rssKb / KB_PER_MB).toFixed(0)} MB`;
    } catch {
        // Not Linux: the line keeps "unavailable".
    }

    return `rss ${(rss / BYTES_PER_MB).toFixed(0)} MB, heap ${(heapUsed / BYTES_PER_MB).toFixed(0)} MB, ${browser}`;
}

function latencyLine(latencies: number[], seconds: number): string {
    const [first, ...rest] = latencies.toSorted((a, b) => a - b);

    if (first === undefined) {
        return "0 renders";
    }

    const sorted: [number, ...number[]] = [first, ...rest];

    return `${sorted.length} renders, ${(sorted.length / seconds).toFixed(2)}/s, p50 ${percentile(sorted, 0.5).toFixed(0)} ms, p95 ${percentile(sorted, 0.95).toFixed(0)} ms, p99 ${percentile(sorted, 0.99).toFixed(0)} ms`;
}

if (import.meta.main) {
    const config = benchLoadConfig();
    const page = PAGE_SIZES.SCREEN_16X9.landscape;
    const png = readFileSync(join(import.meta.dir, "../tests/visual/wasm/images/badge.png")).subarray(0, BROKEN_PNG_BYTES);
    const brokenImage = `data:image/png;base64,${Buffer.from(png).toString("base64")}`;
    // The mix: text, tables with pagination, images, emoji and East Asian text, and one request in six that must fail.
    const mix = [DECKS.text, DECKS.tables, DECKS.images, DECKS.emoji, DECKS.shapes];

    console.log(`load: engines ${config.engines.join(",")}, concurrency ${config.concurrency.join(",")}, ${config.seconds}s each, sample every ${config.sampleSeconds}s`);
    console.log(`start: ${memoryLine()}`);

    for (const engine of config.engines) {
        // Puppeteer: the Node entry's default, a renderer per deck over the shared browser. WASM and PDFKit: one renderer
        // for every request (wasm-pool: createWasmRendererPool() with its default worker count), with system fonts loaded once and unsupported objects reported, as tests/visual/wasm/backend.ts.
        let renderer: PresentationRenderer | undefined;

        if (engine === "wasm") {
            const { createRenderer } = await import("../tests/visual/wasm/backend");

            renderer = await createRenderer();
        } else if (engine === "wasm-pool") {
            const { createWasmRendererPool, loadSystemFonts } = await import("../packages/wasm/src/node");

            renderer = await createWasmRendererPool({ fonts: await loadSystemFonts({ dirs: process.env.FONT_DIRS?.split(":") }), onUnsupported: "warn" });
        } else if (engine === "pdfkit") {
            const { createPdfkitRenderer, loadSystemFonts } = await import("../packages/pdfkit/src/node");

            renderer = await createPdfkitRenderer({ fonts: await loadSystemFonts({ dirs: process.env.FONT_DIRS?.split(":") }), onUnsupported: "warn" });
        }

        console.log(`${engine} ready: ${memoryLine()}`);

        for (const concurrency of config.concurrency) {
            const all: number[] = [];
            let window: number[] = [];
            let request = 0;
            let expectedFailures = 0;
            const errors = new Map<string, number>();
            const start = performance.now();
            const deadline = start + config.seconds * MS_PER_S;
            let windowStart = start;
            const sampler = setInterval(() => {
                const now = performance.now();

                console.log(`${engine} x${concurrency} t=${((now - start) / MS_PER_S).toFixed(0)}s: ${latencyLine(window, (now - windowStart) / MS_PER_S)}; ${memoryLine()}`);
                window = [];
                windowStart = now;
            }, config.sampleSeconds * MS_PER_S);
            const workers: Promise<void>[] = [];

            for (let worker = 0; worker < concurrency; worker++) {
                workers.push(
                    (async () => {
                        while (performance.now() < deadline) {
                            const index = request++;
                            const deck = mix[index % (mix.length + 1)];
                            const presentation = new PuppeteerGen(page, { renderer });

                            if (deck === undefined) {
                                presentation.addSlide().addImage({ data: brokenImage, x: 1, y: 1, w: 2, h: 2 });
                            } else {
                                deck.populate(presentation);
                            }

                            const begin = performance.now();

                            try {
                                await presentation.toPdf();
                                const ms = performance.now() - begin;

                                all.push(ms);
                                window.push(ms);
                                if (deck === undefined) {
                                    errors.set(BROKEN_RENDERED, (errors.get(BROKEN_RENDERED) ?? 0) + 1);
                                }
                            } catch (error) {
                                const message = error instanceof Error ? error.message : String(error);

                                if (deck === undefined) {
                                    expectedFailures++;
                                } else {
                                    errors.set(message, (errors.get(message) ?? 0) + 1);
                                }
                            }
                        }
                    })(),
                );
            }

            await Promise.all(workers);
            clearInterval(sampler);
            const elapsed = (performance.now() - start) / MS_PER_S;

            console.log(`${engine} x${concurrency} total: ${latencyLine(all, elapsed)}, ${expectedFailures} broken-image requests failed as expected, ${errors.size === 0 ? "no other errors" : "other errors:"}`);
            for (const [message, count] of errors) {
                console.log(`  ${count} x ${message}`);
            }

            // After the load: the idle browser closes, then a full collection shows what memory stays.
            await new Promise(resolve => setTimeout(resolve, SETTLE_MS));
            Bun.gc(true);
            console.log(`${engine} x${concurrency} after ${SETTLE_MS / MS_PER_S}s idle: ${memoryLine()}`);
        }
    }

    process.exit(0);
}
