// The seed and run counts of the fast-check properties (tests/*.fuzz.test.ts). Fixed on every `bun test`, so a PR
// sees the same cases; the nightly fuzz job (.github/workflows/fuzz.yml) sets a new seed and more runs. A failure
// prints the seed: rerun with FUZZ_SEED=<seed> to reproduce it.

/** The environment variable that replaces the fixed seed. */
export const FUZZ_SEED_ENV = "FUZZ_SEED";
/** The environment variable that multiplies every property's run count. */
export const FUZZ_SCALE_ENV = "FUZZ_SCALE";

const FIXED_SEED = 661;

function integer(name: string, fallback: number, least: number): number {
    const raw = process.env[name];

    if (raw === undefined || raw === "") {
        return fallback;
    }

    const value = Number(raw);

    if (!Number.isSafeInteger(value) || value < least) {
        throw new Error(`${name} must be an integer of at least ${least}, got ${JSON.stringify(raw)}`);
    }

    return value;
}

/** The seed every property runs with. */
export const FUZZ_SEED = integer(FUZZ_SEED_ENV, FIXED_SEED, Number.MIN_SAFE_INTEGER);

/** `runs`, a property's run count on every `bun test`, times FUZZ_SCALE. */
export function fuzzRuns(runs: number): number {
    return runs * integer(FUZZ_SCALE_ENV, 1, 1);
}

/** `ms`, a property's test timeout on every `bun test`, times FUZZ_SCALE, so more runs get as long per run. */
export function fuzzTimeout(ms: number): number {
    return ms * integer(FUZZ_SCALE_ENV, 1, 1);
}
