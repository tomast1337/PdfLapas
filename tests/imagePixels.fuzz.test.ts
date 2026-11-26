// Property-based tests for checkImagePixels() (#661): whatever an attacker writes, the check either reads every image
// or refuses with one of its documented errors, never crashes, hangs or lets an over-limit image through.
import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { checkImagePixels, IMAGE_PIXEL_LIMITS } from "../packages/model/src/index";

// Every property fixes its seed, so a red build names the same counterexample on every machine until it is fixed.
const SEED = 661;

/** The messages checkImagePixels is allowed to refuse with; any other error escaping it is a bug. */
const REFUSALS = [
    /^An SVG image holds images nested more than \d+ levels deep, so their size cannot be checked; flatten it$/,
    /^The presentation's images nest too deeply to check their size; flatten its SVG images$/,
    /^An image's header is (?:truncated|malformed), so its size cannot be checked against the pixel limits; export the image again$/,
    /^An image in (?:ICO|AVIF|JPEG XL) format cannot be checked against the pixel limits; convert it to PNG, JPEG, GIF or WebP$/,
    /^An SVG image in the [\s\S]* encoding cannot be checked for the images it holds; save it as UTF-8$/,
    /^An image of \d+x\d+ pixels is over the limit of \d+ pixels per image; downscale it, or raise the renderer option imageLimits\.maxImagePixels$/,
    /^An image of \d+x\d+ pixels takes the presentation's images over the limit of \d+ pixels in total; downscale its images, or raise the renderer option imageLimits\.maxDeckImagePixels$/,
];

/** Fails unless `html` is checked or refused with a documented message: a TypeError, RangeError or DOMException is a bug. */
function checkOnlyDocumentedRefusals(html: string): void {
    try {
        checkImagePixels(html, IMAGE_PIXEL_LIMITS);
    } catch (error) {
        if (!(error instanceof Error) || !REFUSALS.some(pattern => pattern.test(error.message))) {
            throw error;
        }
    }
}

// Writing a URI: the scheme's case, the media type and its parameters, a base64 or percent-encoded payload, character
// references in the scheme or the payload, whitespace in the base64, and CSS escapes inside a url().
type Scheme = "data" | "DATA" | "Data" | "dAtA";
type Encoding = "base64" | "percent";
type References = "none" | "scheme" | "payload";
type Writing = {
    readonly scheme: Scheme;
    readonly media: string;
    readonly encoding: Encoding;
    readonly references: References;
    readonly spacing: boolean;
    readonly style: boolean;
};

const SCHEMES: readonly Scheme[] = ["data", "DATA", "Data", "dAtA"];
const ENCODINGS: readonly Encoding[] = ["base64", "percent"];
const REFERENCES: readonly References[] = ["none", "scheme", "payload"];
// The media type of a raster image the check reads by its bytes: Chromium sniffs them, so the type may be anything.
const RASTER_MEDIA = ["image/png", "IMAGE/PNG", "image/png;charset=utf-8", "application/octet-stream", "image/png;x-foo=bar"];
// An image the check scans as an SVG document, however its type is cased and whatever extra parameters it carries.
const SVG_MEDIA = ["image/svg+xml", "IMAGE/SVG+XML", "image/svg+xml;charset=utf-8", "image/svg+xml;x-foo=bar"];

/** Every character of `text` as a decimal HTML character reference, which an attribute decodes before the check reads it. */
function referenceEveryCharacter(text: string): string {
    let out = "";

    for (const character of text) {
        out += `&#${character.charCodeAt(0)};`;
    }

    return out;
}

/** The scheme as character references: the letters decimal, the colon named, as an attribute may spell them. */
function referenceScheme(scheme: Scheme): string {
    let out = "";

    for (const character of `${scheme}:`) {
        out += character === ":" ? "&colon;" : `&#${character.charCodeAt(0)};`;
    }

    return out;
}

/** CSS escapes every non-alphanumeric character as `\hh `, which CSS decodes back; alphanumerics need no escape. */
function cssEscape(text: string): string {
    let out = "";

    for (const character of text) {
        out += /[A-Za-z0-9]/.test(character) ? character : `\\${character.charCodeAt(0).toString(16)} `;
    }

    return out;
}

/** Base64 with ASCII whitespace between its characters: a decoder skips it, so the payload still decodes. */
function insertWhitespace(text: string): string {
    let out = "";

    for (let index = 0; index < text.length; index++) {
        out += text[index] ?? "";
        if (index % 16 === 15) {
            out += " \n\t";
        }
    }

    return out;
}

function base64(bytes: readonly number[]): string {
    return Buffer.from(bytes).toString("base64");
}

function percentEncoded(bytes: readonly number[]): string {
    let out = "";

    for (const byte of bytes) {
        out += `%${byte.toString(16).padStart(2, "0")}`;
    }

    return out;
}

/**
 * The data URI for `bytes` written as `writing` asks, and the markup that carries it: an image `src` or a CSS
 * `url()` with its escapes. Only the literal scheme, media type and comma are CSS-escaped when character references
 * write the payload, since escaping the references' own `;` would put a stray `;` inside the decoded payload.
 */
function markupFor(bytes: readonly number[], writing: Writing): string {
    let encoded = writing.encoding === "base64" ? base64(bytes) : percentEncoded(bytes);

    if (writing.spacing && writing.encoding === "base64") {
        encoded = insertWhitespace(encoded);
    }

    let payload = writing.references === "payload" ? referenceEveryCharacter(encoded) : encoded;
    const media = `${writing.media}${writing.encoding === "base64" ? ";base64" : ""}`;
    const literalScheme = writing.references === "scheme" ? referenceScheme(writing.scheme) : `${writing.scheme}:`;

    if (writing.style) {
        if (writing.references !== "payload") {
            payload = cssEscape(payload);
        }

        const scheme = writing.references === "scheme" ? literalScheme : cssEscape(literalScheme);

        return `<div style="background:url(${scheme}${cssEscape(media)}\\2c ${payload})"></div>`;
    }

    return `<img src="${literalScheme}${media},${payload}">`;
}

const be32 = (value: number) => [value >>> 24, (value >>> 16) & 0xff, (value >>> 8) & 0xff, value & 0xff];

/** A PNG that is only its signature and an RGBA IHDR naming `width` x `height`: no pixel data, so only the header is read. */
function headerOnlyPng(width: number, height: number): number[] {
    return [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, ...be32(13), 0x49, 0x48, 0x44, 0x52, ...be32(width), ...be32(height), 8, 6, 0, 0, 0, 0, 0, 0, 0];
}

/** An SVG document of `width` x `height` holding one PNG image, as UTF-8 bytes Chromium's XML parser reads. */
function svgWithImage(image: string, width: number, height: number): number[] {
    return [...new TextEncoder().encode(`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}"><image href="${image}" width="${width}" height="${height}"/></svg>`)];
}

// 9000 x 9000 is 81,000,000 pixels, over the per-image limit of 67,108,864; 512 x 512 is far under it.
const OVER = 9000;
const UNDER = 512;
const OVER_LIMIT = `is over the limit of ${IMAGE_PIXEL_LIMITS.maxImagePixels} pixels per image`;
const hugePng = headerOnlyPng(OVER, OVER);
const hugeSvg = svgWithImage(`data:image/png;base64,${base64(hugePng)}`, OVER, OVER);
const smallPng = headerOnlyPng(UNDER, UNDER);
const smallSvg = svgWithImage(`data:image/png;base64,${base64(smallPng)}`, UNDER, UNDER);

const schemeArbitrary: fc.Arbitrary<Scheme> = fc.constantFrom(...SCHEMES);
const encodingArbitrary: fc.Arbitrary<Encoding> = fc.constantFrom(...ENCODINGS);
const referencesArbitrary: fc.Arbitrary<References> = fc.constantFrom(...REFERENCES);
const rasterMediaArbitrary: fc.Arbitrary<string> = fc.constantFrom(...RASTER_MEDIA);
const svgMediaArbitrary: fc.Arbitrary<string> = fc.constantFrom(...SVG_MEDIA);

const writingCoreArbitrary = fc.record({
    scheme: schemeArbitrary,
    encoding: encodingArbitrary,
    references: referencesArbitrary,
    spacing: fc.boolean(),
    style: fc.boolean(),
});

// The characters a data URI and its encodings are spelled with, so generated HTML reaches the scanner's branches.
const URI_CHARACTERS = [
    "d",
    "a",
    "t",
    "A",
    "D",
    "T",
    ":",
    ",",
    ";",
    "=",
    "%",
    "&",
    "#",
    "x",
    "X",
    "\\",
    " ",
    "\t",
    "\n",
    "i",
    "m",
    "g",
    "e",
    "/",
    "+",
    "0",
    "9",
    "q",
    "Z",
    "'",
    '"',
    "<",
    ">",
    "u",
    "l",
    "r",
    "(",
    ")",
    "-",
    "_",
    "b",
    "c",
    "s",
    "v",
    "\u0000",
    "é",
    "\u{1f431}",
];
const uriTextArbitrary = fc.string({ unit: fc.constantFrom(...URI_CHARACTERS), maxLength: 200, size: "max" });
/** A few hundred KB of text, as the payload of the raw data URI the timing property prepends. */
const largeTextArbitrary = fc.string({ unit: fc.constantFrom(...URI_CHARACTERS), minLength: 100_000, maxLength: 200_000 });
/** Pieces that spell data URIs and their encodings at every position the scanner looks, appended to the text. */
const largeHtmlPieceArbitrary = fc.oneof(
    fc.constant("data:"),
    fc.constant("DATA:"),
    fc.constant("&colon;"),
    fc.constant("\\3a "),
    fc.constant(","),
    fc.constant(";base64,"),
    fc.constant(")"),
    fc.string({ unit: fc.constantFrom(...URI_CHARACTERS), maxLength: 48 }),
);
const largeHtmlPiecesArbitrary = fc.array(largeHtmlPieceArbitrary, { maxLength: 2000, size: "max" });

describe("checkImagePixels fuzz (#661)", () => {
    test("arbitrary strings end in a documented refusal or in no error at all", () => {
        fc.assert(
            fc.property(fc.string({ unit: "binary", maxLength: 2000, size: "max" }), html => {
                checkOnlyDocumentedRefusals(html);
            }),
            { seed: SEED, numRuns: 500 },
        );
    });

    test("fragments and data URIs with arbitrary meta and payload end in a documented refusal or in no error", () => {
        fc.assert(
            fc.property(fc.string({ unit: "binary", maxLength: 100 }), uriTextArbitrary, uriTextArbitrary, fc.boolean(), (fragment, meta, payload, inAttribute) => {
                const html = inAttribute ? `<div>${fragment}</div><img src="data:${meta},${payload}"><div style="background:url(${fragment}data:${meta},${payload})">${fragment}</div>` : `${fragment}data:${meta},${payload}${fragment}`;

                checkOnlyDocumentedRefusals(html);
            }),
            { seed: SEED, numRuns: 1000 },
        );
    });

    test("the time a check takes stays proportional to the input it reads", () => {
        fc.assert(
            fc.property(largeTextArbitrary, largeHtmlPiecesArbitrary, (text, pieces) => {
                // A raw data URI whose payload is the large text, followed by pieces that spell more URIs and their
                // encodings, so both the bulk scan and the payload decoding see a few hundred KB.
                let html = `data:image/png,${text}`;

                for (const piece of pieces) {
                    html += piece;
                }

                const started = performance.now();

                checkOnlyDocumentedRefusals(html);

                const elapsed = performance.now() - started;

                // The scan charges each view of the text to a budget of 64 characters per input character, so a few
                // hundred KB reads at most a few dozen MB of characters. A measured 300 KB adversarial input takes
                // under 100 ms, so 1 s leaves an order of magnitude for a slow CI machine without hiding a quadratic
                // scan, which on this input would take minutes.
                expect(elapsed).toBeLessThan(1000);
            }),
            { seed: SEED, numRuns: 20 },
        );
    });

    test("an image over the limit is refused however its URI is written", () => {
        fc.assert(
            fc.property(rasterMediaArbitrary, svgMediaArbitrary, writingCoreArbitrary, (rasterMedia, svgMedia, core) => {
                const rasterWriting: Writing = { ...core, media: rasterMedia };
                const svgWriting: Writing = { ...core, media: svgMedia };

                expect(() => checkImagePixels(markupFor(hugePng, rasterWriting), IMAGE_PIXEL_LIMITS)).toThrow(OVER_LIMIT);
                expect(() => checkImagePixels(markupFor(hugeSvg, svgWriting), IMAGE_PIXEL_LIMITS)).toThrow(OVER_LIMIT);
            }),
            { seed: SEED, numRuns: 500 },
        );
    });

    test("the same images under the limit are accepted however their URI is written", () => {
        fc.assert(
            fc.property(rasterMediaArbitrary, svgMediaArbitrary, writingCoreArbitrary, (rasterMedia, svgMedia, core) => {
                const rasterWriting: Writing = { ...core, media: rasterMedia };
                const svgWriting: Writing = { ...core, media: svgMedia };

                expect(() => checkImagePixels(markupFor(smallPng, rasterWriting), IMAGE_PIXEL_LIMITS)).not.toThrow();
                expect(() => checkImagePixels(markupFor(smallSvg, svgWriting), IMAGE_PIXEL_LIMITS)).not.toThrow();
            }),
            { seed: SEED, numRuns: 500 },
        );
    });
});
