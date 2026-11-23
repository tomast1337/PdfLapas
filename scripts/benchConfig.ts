// The bench environment, named once for scripts/bench.ts and scripts/bench-renderers.ts.

/** Runs per measurement. Each bench has its own default: a Puppeteer deck takes far longer than a warm WASM or PDFKit render. */
export const BENCH_RUNS_DEFAULT = { puppeteer: 5, rendererWarm: 20 } as const;

/** `BENCH_RUNS` when set, else the bench's default. */
export function benchRuns(bench: keyof typeof BENCH_RUNS_DEFAULT): number {
    const value = process.env.BENCH_RUNS;

    if (value === undefined) {
        return BENCH_RUNS_DEFAULT[bench];
    }

    const runs = Number(value);

    if (!Number.isInteger(runs) || runs < 1) {
        throw new Error(`BENCH_RUNS must be a positive integer, got "${value}"`);
    }

    return runs;
}

/** The renderers `bun run bench:load` can drive. */
export const LOAD_ENGINES = ["puppeteer", "wasm", "wasm-pool", "wasm-pool-html", "pdfkit", "pdfkit-pool", "native", "native-pool"] as const;

export type LoadEngine = (typeof LOAD_ENGINES)[number];

/** `bun run bench:load` settings, read from the environment variables of LOAD_DEFAULTS. */
export type LoadConfig = { readonly engines: readonly LoadEngine[]; readonly concurrency: readonly number[]; readonly seconds: number; readonly sampleSeconds: number };

/** The load test's environment variables and their defaults. */
export const LOAD_DEFAULTS = { BENCH_LOAD_ENGINES: "puppeteer,wasm", BENCH_LOAD_CONCURRENCY: "1,4,16", BENCH_LOAD_SECONDS: "60", BENCH_LOAD_SAMPLE_SECONDS: "10" } as const;

type LoadVariable = keyof typeof LOAD_DEFAULTS;

/** Reads the load test's environment: comma-separated engines and concurrency levels, seconds per level and per sample. */
export function benchLoadConfig(env: Readonly<Record<string, string | undefined>> = process.env): LoadConfig {
    const read = (name: LoadVariable) => env[name] ?? LOAD_DEFAULTS[name];
    const positive = (name: LoadVariable, text: string) => {
        const value = Number(text);

        if (!Number.isInteger(value) || value < 1) {
            throw new Error(`${name} must hold positive integers, got "${read(name)}"`);
        }

        return value;
    };
    const engines: LoadEngine[] = [];

    for (const name of read("BENCH_LOAD_ENGINES").split(",")) {
        const engine = LOAD_ENGINES.find(known => known === name.trim());

        if (engine === undefined) {
            throw new Error(`BENCH_LOAD_ENGINES lists "${name}"; use a comma-separated list of ${LOAD_ENGINES.join(", ")}`);
        }

        engines.push(engine);
    }

    const concurrency: number[] = [];

    for (const level of read("BENCH_LOAD_CONCURRENCY").split(",")) {
        concurrency.push(positive("BENCH_LOAD_CONCURRENCY", level.trim()));
    }

    return { engines, concurrency, seconds: positive("BENCH_LOAD_SECONDS", read("BENCH_LOAD_SECONDS")), sampleSeconds: positive("BENCH_LOAD_SAMPLE_SECONDS", read("BENCH_LOAD_SAMPLE_SECONDS")) };
}
