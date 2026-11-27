// Property-based tests for the bytes the PDFKit renderer decodes in JavaScript (#661, fifth piece): font files passed to
// createPdfkitRenderer({ fonts }) (fontkit) and images passed to addImage() (PDFKit's PNG and JPEG readers, the GIF and
// WebP decoders). Mutated bytes either work or are refused with an error that is not a JavaScript-engine error: an
// engine error means a damaged file reached code that did not check it, as the corrupt PNG of #659 did.
import { beforeAll, describe, expect, test } from "bun:test";
import { readdirSync, readFileSync } from "node:fs";
import { extname, join } from "node:path";
import fc from "fast-check";
import { create } from "fontkit";
import { createPdfkitRenderer, loadSystemFonts } from "../packages/pdfkit/src/node";
import { PuppeteerGen } from "../packages/pptxgenjs/src/html";
import { FUZZ_SEED, fuzzRuns, fuzzTimeout } from "./fuzzConfig";

const IMAGES = join(import.meta.dir, "visual/wasm/images");
const MIME: Readonly<Record<string, string>> = { ".png": "image/png", ".jpg": "image/jpeg", ".gif": "image/gif", ".webp": "image/webp" };
// One TrueType (glyf) and one CFF face, small enough that a run parses it quickly.
const BASE_FONTS = /^(?:LiberationSans-Regular\.ttf|DejaVuSans\.ttf|NimbusSans-Regular\.otf|NimbusSansNarrow-Bold\.otf)$/;
const TEXT = "Hello Wörld 123 fiffl ÆØ €";

/** Messages the JavaScript engine raises for a value it did not expect: an internal failure, not a documented refusal. */
const INTERNAL_FAILURE = [
    /Cannot read propert/,
    /Cannot set propert/,
    /is not a function/,
    /is not iterable/,
    /is not an object/,
    /Cannot convert undefined or null/,
    /Cannot destructure/,
    /Cannot use 'in' operator/,
    /Invalid array length/,
    /Invalid string length/,
    /Invalid typed array length/,
    /Maximum call stack size exceeded/,
    /out of memory/i,
    /Out of bounds access/,
    /Offset is outside the bounds/,
    /Array buffer allocation failed/,
    /^unsupported number/,
];

/** Fails when `error` is a JavaScript-engine error (a crash); a refusal that says what is wrong is accepted. */
function expectRefusal(error: unknown): void {
    if (!(error instanceof Error)) {
        throw new Error(`a damaged file threw a value that is not an Error: ${String(error)}`);
    }

    // A refusal that wraps the decoder's error (`cause`) quotes its message; only the wrapper's own words are checked.
    const own = error.cause instanceof Error ? error.message.replace(error.cause.message, "") : error.message;

    for (const pattern of INTERNAL_FAILURE) {
        if (pattern.test(own)) {
            throw error;
        }
    }
}

// --- Mutating bytes -------------------------------------------------------------------------------------------

type Edit =
    | { readonly kind: "byte"; readonly at: number; readonly value: number }
    | { readonly kind: "word"; readonly at: number; readonly value: number; readonly size: 2 | 4 }
    | { readonly kind: "cut"; readonly at: number }
    | { readonly kind: "repeat"; readonly at: number; readonly length: number; readonly times: number };

// Values that break a length, an offset or a count: zero, one, the largest of each width and the sign bit.
const EXTREMES = [0, 1, 0x7f, 0x80, 0xff, 0x7fff, 0x8000, 0xffff, 0x7fffffff, 0x80000000, 0xffffffff];
const inPlace = [
    { weight: 4, arbitrary: fc.record({ kind: fc.constant("byte" as const), at: fc.nat(), value: fc.integer({ min: 0, max: 255 }) }) },
    { weight: 4, arbitrary: fc.record({ kind: fc.constant("word" as const), at: fc.nat(), value: fc.constantFrom(...EXTREMES), size: fc.constantFrom(2 as const, 4 as const) }) },
    { weight: 1, arbitrary: fc.record({ kind: fc.constant("cut" as const), at: fc.nat() }) },
] as const;
const edit: fc.Arbitrary<Edit> = fc.oneof(...inPlace, { weight: 1, arbitrary: fc.record({ kind: fc.constant("repeat" as const), at: fc.nat(), length: fc.integer({ min: 1, max: 64 }), times: fc.integer({ min: 2, max: 16 }) }) });
// A font's edits leave every table where it was: a repeat would move the layout tables' data under their offsets.
const fontEdit: fc.Arbitrary<Edit> = fc.oneof(...inPlace);

/**
 * `bytes` with each edit applied, at an offset inside one of `regions` (start, end): the parts of the file a decoder
 * reads first (headers, table directories, the tables a render touches), where an edit reaches more code than a
 * random offset in glyph or pixel data would.
 */
function mutate(bytes: Uint8Array, regions: readonly (readonly [number, number])[], edits: readonly Edit[]): Uint8Array {
    let result = Uint8Array.from(bytes);

    for (const [index, change] of edits.entries()) {
        const [start, end] = regions[index % regions.length] ?? [0, result.length];
        const span = Math.max(1, Math.min(end, result.length) - start);
        const at = Math.min(result.length - 1, start + (change.at % span));

        if (at < 0) {
            break;
        }

        if (change.kind === "byte") {
            result[at] = change.value;
        } else if (change.kind === "word") {
            const view = new DataView(result.buffer);

            if (change.size === 2 && at + 2 <= result.length) {
                view.setUint16(at, change.value & 0xffff);
            } else if (change.size === 4 && at + 4 <= result.length) {
                view.setUint32(at, change.value >>> 0);
            }
        } else if (change.kind === "cut") {
            result = result.slice(0, at);
        } else {
            const piece = result.slice(at, at + change.length);
            const repeated = new Uint8Array(result.length + piece.length * (change.times - 1));

            repeated.set(result.subarray(0, at));

            for (let copy = 0; copy < change.times; copy++) {
                repeated.set(piece, at + copy * piece.length);
            }

            repeated.set(result.subarray(at + piece.length), at + piece.length * change.times);
            result = repeated;
        }
    }

    return result;
}

// shortcut: the OpenType layout tables and their table records are not mutated: fontkit 2.0.4 can exhaust memory on a
// damaged one, which nothing in the same process can stop (createPdfkitRendererPool's workerHeapMb contains it under
// Node). Add them back once fontkit reads them safely.
const LAYOUT_TABLES: readonly string[] = ["GSUB", "GPOS", "GDEF"];

/**
 * The byte ranges of an sfnt's header, of its table records but the layout tables', and of the tables PDFKit reads to
 * lay out and embed text.
 */
function fontRegions(bytes: Uint8Array): [number, number][] {
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const count = view.getUint16(4);
    const regions: [number, number][] = [[0, 12]];

    for (let index = 0; index < count; index++) {
        const record = 12 + index * 16;
        const tag = String.fromCharCode(bytes[record] ?? 0, bytes[record + 1] ?? 0, bytes[record + 2] ?? 0, bytes[record + 3] ?? 0);
        const offset = view.getUint32(record + 8);
        const length = view.getUint32(record + 12);

        if (LAYOUT_TABLES.includes(tag)) {
            continue;
        }

        regions.push([record, record + 16]);

        // The glyph tables are large: only their start, where the first glyphs and the CFF header and indexes are.
        if (["head", "hhea", "maxp", "hmtx", "cmap", "loca", "name", "OS/2", "post", "kern"].includes(tag)) {
            regions.push([offset, offset + length]);
        } else if (tag === "glyf" || tag === "CFF ") {
            regions.push([offset, offset + Math.min(length, 4096)]);
        }
    }

    return regions;
}

// --- Properties -----------------------------------------------------------------------------------------------

describe("PDFKit renderer over damaged font and image files (#661)", () => {
    let fonts: { readonly bytes: Uint8Array; readonly family: string; readonly regions: [number, number][] }[] = [];
    let images: { readonly bytes: Uint8Array; readonly mime: string }[] = [];
    let fallback: Uint8Array[] = [];

    beforeAll(async () => {
        const found = await loadSystemFonts({ match: BASE_FONTS });

        for (const bytes of found) {
            // `as`: fontkit reads any Uint8Array; its types name Node's Buffer. The base files are whole fonts, not collections.
            const font = create(bytes as Buffer);

            if (!("fonts" in font)) {
                fonts.push({ bytes, family: font.familyName, regions: fontRegions(bytes) });
            }
        }

        fallback = await loadSystemFonts();

        for (const file of readdirSync(IMAGES)) {
            const mime = MIME[extname(file)];

            if (mime !== undefined) {
                images.push({ bytes: readFileSync(join(IMAGES, file)), mime });
            }
        }
    });

    test("the base files load and render", async () => {
        expect(fonts.length).toBeGreaterThan(0);
        expect(images.length).toBeGreaterThan(0);

        for (const font of fonts) {
            const renderer = await createPdfkitRenderer({ fonts: [font.bytes], onUnsupported: "warn", onMissingFont: "ignore" });
            const deck = new PuppeteerGen();

            deck.addSlide().addText(TEXT, { x: 1, y: 1, w: 6, h: 1, fontFace: font.family, bold: true });
            expect(new TextDecoder().decode((await renderer.render(await deck.toHtml())).subarray(0, 5))).toBe("%PDF-");
        }
    });

    test(
        "a damaged font is refused when the renderer is created or when it draws, never with an engine error",
        async () => {
            await fc.assert(
                fc.asyncProperty(fc.nat(), fc.array(fontEdit, { minLength: 1, maxLength: 4 }), async (which, edits) => {
                    const font = fonts[which % fonts.length];

                    if (font === undefined) {
                        return;
                    }

                    const deck = new PuppeteerGen();

                    deck.addSlide().addText(
                        [
                            { text: TEXT, options: { fontFace: font.family, breakLine: true } },
                            { text: TEXT, options: { fontFace: font.family, bold: true, italic: true, underline: { style: "sng" }, charSpacing: 2 } },
                        ],
                        { x: 1, y: 1, w: 6, h: 2 },
                    );

                    try {
                        const renderer = await createPdfkitRenderer({ fonts: [mutate(font.bytes, font.regions, edits)], onUnsupported: "warn", onMissingFont: "ignore" });
                        const pdf = await renderer.render(await deck.toHtml());

                        expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe("%PDF-");
                    } catch (error) {
                        expectRefusal(error);
                    }
                }),
                { seed: FUZZ_SEED, numRuns: fuzzRuns(150) },
            );
        },
        fuzzTimeout(120_000),
    );

    test(
        "a damaged image is refused when it is added, checked or drawn, never with an engine error",
        async () => {
            const renderer = await createPdfkitRenderer({ fonts: fallback, onUnsupported: "warn", onMissingFont: "ignore" });

            await fc.assert(
                fc.asyncProperty(fc.nat(), fc.array(edit, { minLength: 1, maxLength: 4 }), fc.boolean(), async (which, edits, header) => {
                    const image = images[which % images.length];

                    if (image === undefined) {
                        return;
                    }

                    // Half the runs edit the first 64 bytes only: the signature, the header and the first chunk or segment.
                    const bytes = mutate(image.bytes, [header ? [0, 64] : [0, image.bytes.length]], edits);
                    const deck = new PuppeteerGen();

                    try {
                        deck.addSlide().addImage({ data: `data:${image.mime};base64,${Buffer.from(bytes).toString("base64")}`, x: 1, y: 1, w: 2, h: 2 });

                        const pdf = await renderer.render(await deck.toHtml());

                        expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe("%PDF-");
                    } catch (error) {
                        expectRefusal(error);
                    }
                }),
                { seed: FUZZ_SEED, numRuns: fuzzRuns(300) },
            );
        },
        fuzzTimeout(120_000),
    );
});
