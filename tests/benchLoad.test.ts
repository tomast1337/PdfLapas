import { describe, expect, test } from "bun:test";
import { descendants, percentile } from "../scripts/bench-load";
import { benchLoadConfig, LOAD_DEFAULTS } from "../scripts/benchConfig";

describe("bench:load", () => {
    test("percentile is the nearest rank", () => {
        const sorted: [number, ...number[]] = [10, 20, 30, 40, 50, 60, 70, 80, 90, 100];

        expect(percentile(sorted, 0.5)).toBe(50);
        expect(percentile(sorted, 0.95)).toBe(100);
        expect(percentile(sorted, 0)).toBe(10);
        expect(percentile([7], 0.99)).toBe(7);
    });

    test("descendants walks every level below the root and nothing beside it", () => {
        // 1 -> 2 -> 3 -> 4, 1 -> 5, and 6 under another parent.
        const parents = new Map([
            [2, 1],
            [3, 2],
            [4, 3],
            [5, 1],
            [6, 9],
        ]);

        expect(descendants(parents, 1).toSorted()).toEqual([2, 3, 4, 5]);
        expect(descendants(parents, 4)).toEqual([]);
    });

    test("the configuration reads its environment variables and rejects bad values", () => {
        expect(benchLoadConfig({})).toEqual({ engines: ["puppeteer", "wasm"], concurrency: [1, 4, 16], seconds: Number(LOAD_DEFAULTS.BENCH_LOAD_SECONDS), sampleSeconds: Number(LOAD_DEFAULTS.BENCH_LOAD_SAMPLE_SECONDS) });
        expect(benchLoadConfig({ BENCH_LOAD_ENGINES: "wasm, pdfkit", BENCH_LOAD_CONCURRENCY: "2, 64", BENCH_LOAD_SECONDS: "5", BENCH_LOAD_SAMPLE_SECONDS: "1" })).toEqual({ engines: ["wasm", "pdfkit"], concurrency: [2, 64], seconds: 5, sampleSeconds: 1 });
        expect(() => benchLoadConfig({ BENCH_LOAD_ENGINES: "chrome" })).toThrow(/BENCH_LOAD_ENGINES lists "chrome"/);
        expect(() => benchLoadConfig({ BENCH_LOAD_CONCURRENCY: "4,0" })).toThrow(/BENCH_LOAD_CONCURRENCY/);
        expect(() => benchLoadConfig({ BENCH_LOAD_SECONDS: "1.5" })).toThrow(/BENCH_LOAD_SECONDS/);
    });
});
