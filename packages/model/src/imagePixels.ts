// Decoded-pixel limits on the raster images of a presentation. A few KB of compressed data can declare billions of
// pixels, and decoding them exhausts the renderer's memory or time, so every renderer reads each image's size from
// its header and refuses the same decks before decoding. crates/pdf-core/src/image.rs applies the same defaults.

/** The default limits: 8192 x 8192 pixels per image (256 MiB as RGBA) and four times that per presentation. */
export const IMAGE_PIXEL_LIMITS = Object.freeze({
    /** Most pixels one raster image may have. */
    maxImagePixels: 2 ** 26,
    /** Most pixels the distinct raster images of one presentation may have together. */
    maxDeckImagePixels: 2 ** 28,
});

/** Overrides of `IMAGE_PIXEL_LIMITS` (the Puppeteer and PDFKit renderers' `imageLimits` option). */
export type ImagePixelLimits = {
    readonly maxImagePixels?: number;
    readonly maxDeckImagePixels?: number;
};

/** A header field: its value, or `short` when the bytes end before it does. */
type HeaderField = { kind: "value"; value: number } | { kind: "short" };

const be = (bytes: Uint8Array, at: number, count: number): HeaderField => {
    if (at + count > bytes.length) {
        return { kind: "short" };
    }

    let value = 0;

    for (let index = 0; index < count; index++) {
        value = value * 256 + (bytes[at + index] ?? 0);
    }

    return { kind: "value", value };
};
const le = (bytes: Uint8Array, at: number, count: number): HeaderField => {
    if (at + count > bytes.length) {
        return { kind: "short" };
    }

    let value = 0;

    for (let index = count - 1; index >= 0; index--) {
        value = value * 256 + (bytes[at + index] ?? 0);
    }

    return { kind: "value", value };
};
const startsWith = (bytes: Uint8Array, signature: readonly number[], at = 0) => signature.every((byte, index) => bytes[at + index] === byte);

/**
 * The pixel size an image header declares: `size`; `unknown-format` when no raster signature matches (SVG, another
 * format); `truncated` when the bytes end before the size does; `malformed` when the header is not laid out as its
 * format requires (a check on a signature or marker past the end of the bytes also lands here).
 */
export type ImageHeaderSize = { kind: "size"; width: number; height: number } | { kind: "unknown-format" } | { kind: "truncated" } | { kind: "malformed" };

const TRUNCATED = { kind: "truncated" } as const;
const MALFORMED = { kind: "malformed" } as const;
const headerSize = (width: HeaderField, height: HeaderField): ImageHeaderSize => (width.kind === "short" || height.kind === "short" ? TRUNCATED : { kind: "size", width: width.value, height: height.value });

/** WebP: the first chunk is VP8X (canvas size minus one), VP8L (14-bit sizes minus one) or VP8 (14-bit sizes). */
export function webpSize(bytes: Uint8Array): ImageHeaderSize {
    const chunk = String.fromCharCode(...bytes.subarray(12, 16));

    if (chunk === "VP8X") {
        const [w, h] = [le(bytes, 24, 3), le(bytes, 27, 3)];

        return w.kind === "short" || h.kind === "short" ? TRUNCATED : { kind: "size", width: w.value + 1, height: h.value + 1 };
    }

    if (chunk === "VP8L" && bytes[20] === 0x2f) {
        const bits = le(bytes, 21, 4);

        return bits.kind === "short" ? TRUNCATED : { kind: "size", width: (bits.value % 2 ** 14) + 1, height: (Math.floor(bits.value / 2 ** 14) % 2 ** 14) + 1 };
    }

    if (chunk === "VP8 " && startsWith(bytes, [0x9d, 0x01, 0x2a], 23)) {
        const [w, h] = [le(bytes, 26, 2), le(bytes, 28, 2)];

        return w.kind === "short" || h.kind === "short" ? TRUNCATED : { kind: "size", width: w.value % 2 ** 14, height: h.value % 2 ** 14 };
    }

    return MALFORMED;
}

/** The pixel size a raster image declares in its header (PNG IHDR, JPEG SOFn, GIF logical screen, WebP, BMP), by its magic bytes. */
export function rasterImageSize(bytes: Uint8Array): ImageHeaderSize {
    if (startsWith(bytes, [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])) {
        if (!startsWith(bytes, [0x49, 0x48, 0x44, 0x52], 12)) {
            return MALFORMED;
        } // "IHDR"

        return headerSize(be(bytes, 16, 4), be(bytes, 20, 4));
    }

    if (startsWith(bytes, [0xff, 0xd8])) {
        // JPEG: the first start-of-frame segment (SOF0-SOF15 except DHT, JPG and DAC) holds the height, then the width.
        for (let at = 2; at < bytes.length;) {
            if (bytes[at] !== 0xff) {
                return MALFORMED;
            }

            const marker = bytes[at + 1];

            if (marker === undefined) {
                return TRUNCATED;
            }

            if (marker === 0xff) {
                at += 1;
            } else if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) {
                at += 2;
            } else if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
                const [height, width] = [be(bytes, at + 5, 2), be(bytes, at + 7, 2)];

                return headerSize(width, height);
            } else if (marker === 0xda || marker === 0xd9) {
                // The scan or the end of the image before any frame header.
                return MALFORMED;
            } else {
                const length = be(bytes, at + 2, 2);

                if (length.kind === "short") {
                    return TRUNCATED;
                }

                at += 2 + length.value;
            }
        }

        return TRUNCATED;
    }

    if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) {
        return headerSize(le(bytes, 6, 2), le(bytes, 8, 2));
    }

    if (startsWith(bytes, [0x52, 0x49, 0x46, 0x46]) && startsWith(bytes, [0x57, 0x45, 0x42, 0x50], 8)) {
        return webpSize(bytes);
    }

    if (startsWith(bytes, [0x42, 0x4d])) {
        // BMP: a 12-byte core header holds 16-bit sizes, every later header signed 32-bit ones (negative: top-down).
        const header = le(bytes, 14, 4);

        if (header.kind === "short") {
            return TRUNCATED;
        }

        if (header.value === 12) {
            return headerSize(le(bytes, 18, 2), le(bytes, 20, 2));
        }

        const [w, h] = [le(bytes, 18, 4), le(bytes, 22, 4)];

        return w.kind === "short" || h.kind === "short" ? TRUNCATED : { kind: "size", width: Math.abs(w.value | 0), height: Math.abs(h.value | 0) };
    }

    return { kind: "unknown-format" };
}

// Where a data URI starts and its media type and parameters, whatever the type: Chromium sniffs a raster image's bytes,
// so `data:application/octet-stream` draws a PNG as `data:image/png` does.
const DATA_URI = /data:((?:(?!data:)[^,])*),/gi;
// A base64 payload ends at the first character that is not base64, `%` (decoded first) or ASCII whitespace (skipped):
// that character ends the URI or makes it invalid, so nothing past it is decoded.
const BASE64_PAYLOAD = /^[A-Za-z0-9+/=%\s]*/;
const NOT_BASE64 = /[^A-Za-z0-9+/]/g;
const HEX_PAIR = /^[\da-fA-F]{2}$/;
// Base64 characters decoded for the header (a multiple of 4): 48 KiB of bytes holds every header but a JPEG's after
// large metadata, which is then decoded whole.
const HEAD_CHARACTERS = 65_536;
// An SVG image can hold images of its own, and those more: deeper than this is refused, as is more scanning than this
// many characters per character of the HTML, so the check never costs more than the render it guards.
const MAX_NESTING = 8;
const SCAN_BUDGET_PER_CHARACTER = 64;
// Character references (HTML and XML attributes, XML text) and CSS escapes can spell `data:` and its payload. Letters
// and digits have no named reference, so the numeric ones and these cover every character of a data URI.
const CHARACTER_REFERENCE = /&(?:#(\d{1,8})|#[xX]([\da-fA-F]{1,8})|([A-Za-z]{2,8}));?/g;
const NAMED_REFERENCES: Readonly<Record<string, string>> = Object.freeze({
    amp: "&",
    AMP: "&",
    lt: "<",
    LT: "<",
    gt: ">",
    GT: ">",
    quot: '"',
    QUOT: '"',
    apos: "'",
    colon: ":",
    comma: ",",
    semi: ";",
    sol: "/",
    bsol: "\\",
    percnt: "%",
    plus: "+",
    equals: "=",
    lpar: "(",
    rpar: ")",
    num: "#",
    excl: "!",
    period: ".",
    quest: "?",
    commat: "@",
    Tab: "\t",
    NewLine: "\n",
});
const CSS_ESCAPE = /\\(?:([\da-fA-F]{1,6})(?:\r\n|[ \t\n\r\f])?|\r\n|[\n\r\f]|([^]))/g;
const UTF8 = new TextEncoder();

const character = (code: number) => (code > 0 && code <= 0x10ffff ? String.fromCodePoint(code) : "�");

/** Bytes as a string of one character per byte (what `atob` returns and `btoa` takes). */
function binaryString(bytes: Uint8Array): string {
    let text = "";

    for (let at = 0; at < bytes.length; at += 8192) {
        text += String.fromCharCode(...bytes.subarray(at, at + 8192));
    }

    return text;
}

/** Base64 decoded as far as it goes: other characters are dropped, and a last group of one character too. */
function base64Bytes(base64: string): Uint8Array {
    const clean = base64.replace(NOT_BASE64, "");
    const whole = clean.length % 4 === 1 ? clean.slice(0, -1) : clean.padEnd(Math.ceil(clean.length / 4) * 4, "=");

    return Uint8Array.from(atob(whole), letter => letter.charCodeAt(0));
}

/** A URL's characters as the bytes they stand for: `%XX` is that byte, any other character its UTF-8. */
function urlBytes(text: string): Uint8Array {
    const bytes: number[] = [];

    for (let at = 0; at < text.length;) {
        const code = text.charCodeAt(at);
        const hex = code === 0x25 ? text.slice(at + 1, at + 3) : "";

        if (HEX_PAIR.test(hex)) {
            bytes.push(Number.parseInt(hex, 16));
            at += 3;
        } else if (code < 0x80) {
            bytes.push(code);
            at += 1;
        } else {
            const point = character(text.codePointAt(at) ?? 0xfffd);

            for (const byte of UTF8.encode(point)) {
                bytes.push(byte);
            }

            at += point.length;
        }
    }

    return Uint8Array.from(bytes);
}

/** What the HTML or an SVG image holds and every check of it so far shares. */
type ImageScan = {
    readonly limits: Required<ImagePixelLimits>;
    /** Each distinct image or SVG document, counted or scanned once. */
    readonly seen: Set<string>;
    total: number;
    /** Characters left to scan. */
    budget: number;
};

/** What a scan found: every image within the limits, or the message of the first refused one. */
type ScanResult = { kind: "checked" } | { kind: "refused"; message: string };

/** Counts every image `text` holds, wherever a data URI can appear in it, and scans the SVG images it holds in turn. */
function scanImages(text: string, depth: number, scan: ImageScan): ScanResult {
    if (depth > MAX_NESTING) {
        return { kind: "refused", message: `An SVG image holds images nested more than ${MAX_NESTING} levels deep, so their size cannot be checked; flatten it` };
    }

    // The text as each context reads it: an attribute decodes character references, CSS its escapes, a style
    // attribute both. Reading every view, no encoding hides a URI.
    const references = text.replace(CHARACTER_REFERENCE, (reference, decimal: string | undefined, hex: string | undefined, name: string | undefined) => {
        if (name !== undefined) {
            return NAMED_REFERENCES[name] ?? reference;
        }

        return character(decimal === undefined ? Number.parseInt(hex ?? "", 16) : Number.parseInt(decimal, 10));
    });
    const unescape = (css: string) => css.replace(CSS_ESCAPE, (_escape, hex: string | undefined, other: string | undefined) => (hex === undefined ? (other ?? "") : character(Number.parseInt(hex, 16))));

    for (const view of new Set([text, references, unescape(text), unescape(references)])) {
        scan.budget -= view.length;
        if (scan.budget < 0) {
            return { kind: "refused", message: "The presentation's images nest too deeply to check their size; flatten its SVG images" };
        }

        const starts = [...view.matchAll(DATA_URI)];

        for (let index = 0; index < starts.length; index++) {
            const start = starts[index];

            if (start === undefined) {
                continue;
            }

            // The payload runs to the next URI: a quote or bracket the scan took for its end could cut it short.
            const payload = view.slice(start.index + start[0].length, starts[index + 1]?.index ?? view.length);
            const meta = start[1] ?? "";
            // How Chromium reads the parameters, as written and percent-decoded: base64 when the last one is `base64`
            // (older versions: any one), an SVG document by its media type, and the SVG's charset.
            let base64 = false;
            let raw = false;
            let svg = false;
            let charset = "utf-8";

            for (const form of new Set([meta, binaryString(urlBytes(meta))])) {
                const parameters = form.split(";");

                svg ||= parameters[0]?.trim().toLowerCase() === "image/svg+xml";
                for (const [position, parameter] of parameters.entries()) {
                    const [name = "", value] = parameter.split("=");
                    const lowered = name.trim().toLowerCase();

                    if (lowered === "base64") {
                        base64 = true;
                        raw ||= position !== parameters.length - 1;
                    } else if (lowered === "charset" && value !== undefined) {
                        charset = value.trim().toLowerCase();
                    }
                }

                raw ||= parameters.at(-1)?.trim().toLowerCase() !== "base64";
            }

            // Each reading of the payload: its key, so an image used twice counts once, its bytes, and whether its end is
            // exact. A base64 payload's is; an unencoded one runs to the next URI, which its own bytes can spell.
            const readings: { key: string; bytes: () => Uint8Array; head: Uint8Array; exact: boolean }[] = [];

            if (base64) {
                const run = BASE64_PAYLOAD.exec(payload)?.[0] ?? "";
                const decoded = run.includes("%") ? binaryString(urlBytes(run)) : run;
                const compact = decoded.replace(NOT_BASE64, "");

                readings.push({ key: `base64:${compact}`, bytes: () => base64Bytes(compact), head: base64Bytes(compact.slice(0, HEAD_CHARACTERS)), exact: true });
            }

            if (raw) {
                readings.push({ key: `url:${payload}`, bytes: () => urlBytes(payload), head: urlBytes(payload.slice(0, 3 * HEAD_CHARACTERS)), exact: false });
            }

            for (const reading of readings) {
                if (scan.seen.has(reading.key)) {
                    continue;
                }

                scan.seen.add(reading.key);
                let size = rasterImageSize(reading.head);

                if (size.kind !== "size" && startsWith(reading.head, [0xff, 0xd8])) {
                    size = rasterImageSize(reading.bytes());
                }

                // A short or broken header draws nothing, and the renderer reports it; unless the scan cut it short.
                if ((size.kind === "truncated" || size.kind === "malformed") && reading.exact) {
                    continue;
                }

                if (size.kind === "truncated" || size.kind === "malformed") {
                    return { kind: "refused", message: `An image's header is ${size.kind}, so its size cannot be checked against the pixel limits; export the image again` };
                }

                if (size.kind === "unknown-format") {
                    // Formats Chromium decodes whose size this check does not read: refused rather than drawn unchecked.
                    const brand = binaryString(reading.head.subarray(8, 12));
                    const format =
                        startsWith(reading.head, [0, 0, 1, 0]) || startsWith(reading.head, [0, 0, 2, 0])
                            ? "ICO"
                            : startsWith(reading.head, [0x66, 0x74, 0x79, 0x70], 4) && ["avif", "avis", "mif1", "msf1"].includes(brand)
                              ? "AVIF"
                              : startsWith(reading.head, [0xff, 0x0a]) || startsWith(reading.head, [0, 0, 0, 0x0c, 0x4a, 0x58, 0x4c, 0x20])
                                ? "JPEG XL"
                                : "";

                    if (format !== "") {
                        return { kind: "refused", message: `An image in ${format} format cannot be checked against the pixel limits; convert it to PNG, JPEG, GIF or WebP` };
                    }

                    if (!svg) {
                        continue;
                    }

                    // An SVG image draws the images it holds: its text, as Chromium's XML parser decodes it (UTF-16
                    // by its byte order mark or a NUL in its first two bytes, else UTF-8 or the declared charset).
                    const bytes = reading.bytes();
                    const utf16 = bytes[0] === 0xff && bytes[1] === 0xfe ? "utf-16le" : bytes[0] === 0xfe && bytes[1] === 0xff ? "utf-16be" : bytes[1] === 0 ? "utf-16le" : bytes[0] === 0 ? "utf-16be" : "";
                    const document = new TextDecoder(utf16 === "" ? "utf-8" : utf16).decode(bytes);
                    const declared = /^\s*<\?xml[^>]*encoding\s*=\s*["']([^"']*)["']/i.exec(document)?.[1]?.toLowerCase() ?? charset;

                    if (utf16 === "" && !/^utf-?8$/.test(declared)) {
                        return { kind: "refused", message: `An SVG image in the ${declared} encoding cannot be checked for the images it holds; save it as UTF-8` };
                    }

                    const nested = scanImages(document, depth + 1, scan);

                    if (nested.kind === "refused") {
                        return nested;
                    }

                    continue;
                }

                const { width: w, height: h } = size;
                const pixels = w * h;

                if (pixels > scan.limits.maxImagePixels) {
                    return { kind: "refused", message: `An image of ${w}x${h} pixels is over the limit of ${scan.limits.maxImagePixels} pixels per image; downscale it, or raise the renderer option imageLimits.maxImagePixels` };
                }

                scan.total += pixels;
                if (scan.total > scan.limits.maxDeckImagePixels) {
                    return {
                        kind: "refused",
                        message: `An image of ${w}x${h} pixels takes the presentation's images over the limit of ${scan.limits.maxDeckImagePixels} pixels in total; downscale its images, or raise the renderer option imageLimits.maxDeckImagePixels`,
                    };
                }
            }
        }
    }

    return { kind: "checked" };
}

function limit(name: keyof ImagePixelLimits, value: number | undefined): number {
    if (value === undefined) {
        return IMAGE_PIXEL_LIMITS[name];
    }

    if (Number.isInteger(value) && value > 0) {
        return value;
    }

    throw new TypeError(`imageLimits.${name} must be a positive integer, got ${JSON.stringify(value)}`);
}

/** Checks `limits` where they enter (a renderer's constructor) and returns them complete. */
export function imagePixelLimits(limits: ImagePixelLimits = {}): Required<ImagePixelLimits> {
    return { maxImagePixels: limit("maxImagePixels", limits.maxImagePixels), maxDeckImagePixels: limit("maxDeckImagePixels", limits.maxDeckImagePixels) };
}

/**
 * Throws when an image in `html` declares more pixels than `limits.maxImagePixels`, or the distinct images together
 * more than `limits.maxDeckImagePixels`, before any of them is decoded. Every data URI counts, whatever its media type
 * or encoding, and so do the images an SVG image holds; an image whose size cannot be read is refused.
 */
export function checkImagePixels(html: string, limits: Required<ImagePixelLimits>): void {
    const result = scanImages(html, 0, { limits, seen: new Set(), total: 0, budget: (html.length + HEAD_CHARACTERS) * SCAN_BUDGET_PER_CHARACTER });

    if (result.kind === "refused") {
        throw new Error(result.message);
    }
}
