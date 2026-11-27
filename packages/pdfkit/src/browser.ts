// Browser-safe PDFKit renderer ("@pdflapas/pdfkit" under the browser/default conditions): pure JavaScript, no
// Chromium and no wasm. Must not load node:* modules, puppeteer or jsdom (enforced by scripts/build.ts).
import { checkImagePixels, type HtmlPresentationDocument, type ImagePixelLimits, imagePixelLimits, type MissingFont, type MissingFontPolicy, missingFontPolicy, type PresentationRenderer, reportMissingFonts } from "@pdflapas/model";
import { renderHtml } from "./document";
import { type Face, parseFonts } from "./fonts";

/** How PDFKit's number() begins the error it throws for a number it cannot write (pdfkit/js/pdfkit.js). */
const PDFKIT_UNSUPPORTED_NUMBER = "unsupported number: ";

export { type MissingFont, MissingFontError, type MissingFontPolicy } from "@pdflapas/model";

export type UnsupportedPolicy = "error" | "warn";

export type PdfkitRendererOptions = {
    /** TTF/OTF/TTC font files, parsed once here. The renderer sees no system fonts; at least one is required. */
    fonts: readonly Uint8Array[];
    /**
     * What to do when the renderer skips something it cannot draw; each skip is one line of `warnings()`.
     * `"error"` (default) throws and lists them; `"warn"` returns the PDF and leaves the list in `warnings()`.
     */
    onUnsupported?: UnsupportedPolicy;
    /**
     * What to do when a text run asks for a face none of `fonts` provides, neither the face nor its metric-compatible
     * alias (Calibri -> Carlito, Arial -> Liberation Sans, ...): `"warn"` (default), `"throw"` (`MissingFontError`),
     * `"ignore"`, or a function receiving the list. Not part of `onUnsupported`: the text is drawn, in a fallback face.
     */
    onMissingFont?: MissingFontPolicy;
    /** Decoded-pixel limits on the deck's raster images, checked from their headers before drawing. Default IMAGE_PIXEL_LIMITS. */
    imageLimits?: ImagePixelLimits;
};

export type PdfkitRendererPoolOptions = PdfkitRendererOptions & {
    /**
     * Worker threads, each with its own renderer and its own parsed copy of `fonts` (the font bytes are shared). Default:
     * one per core but one, at most 8.
     */
    workers?: number;
    /**
     * Each worker's JavaScript heap limit, in MB (Node's `resourceLimits.maxOldGenerationSizeMb`). A render that needs
     * more ends its worker with an error, and the pool replaces the worker, instead of the whole process running out
     * of memory. Default 2048. Bun does not apply worker resource limits: there a render that runs away takes the
     * process with it.
     */
    workerHeapMb?: number;
};

/** What createPdfkitRendererPool() returns (Node/Bun): a renderer whose renders run in worker threads. */
export type PdfkitRendererPool = PresentationRenderer & {
    /** What the last render to finish could not draw; with several in flight, read the error of the one you need instead. */
    warnings(): string[];
    /** The font faces the last render to finish asked for that none of the fonts provides. */
    missingFonts(): MissingFont[];
    /** Terminates the workers; queued and in-flight renders reject. */
    close(): Promise<void>;
    [Symbol.asyncDispose](): Promise<void>;
};

export type SystemFontOptions = {
    /** Directories searched recursively. Defaults to the platform's system and user font directories. */
    dirs?: readonly string[];
    /** Font file basenames to load. Defaults to the metric-compatible fallback set and the Office faces they stand in for. */
    match?: RegExp;
};

/** PDF backend on PDFKit: renders the PdfLapas HTML profile document `toHtml()` returns (docs/HTML_PROFILE.md). */
export class PdfkitRenderer implements PresentationRenderer {
    readonly #faces: readonly Face[];
    readonly #onUnsupported: UnsupportedPolicy;
    readonly #onMissingFont: MissingFontPolicy;
    readonly #imageLimits: Required<ImagePixelLimits>;
    #warnings: readonly string[] = [];
    #missingFonts: readonly MissingFont[] = [];

    // Takes the options, not parsed faces: a face holds a fontkit font, and naming it here would put fontkit's types in
    // the published declarations, which a strict consumer cannot resolve (#354).
    constructor({ fonts, onUnsupported = "error", onMissingFont, imageLimits }: PdfkitRendererOptions) {
        if (!fonts?.length) {
            throw new Error("createPdfkitRenderer() requires at least one font: pass { fonts: [Uint8Array, ...] } (Node/Bun: await loadSystemFonts())");
        }

        if (onUnsupported !== "error" && onUnsupported !== "warn") {
            throw new Error(`onUnsupported must be "error" or "warn", got ${JSON.stringify(onUnsupported)}`);
        }

        this.#faces = parseFonts(fonts);
        this.#onUnsupported = onUnsupported;
        this.#onMissingFont = missingFontPolicy(onMissingFont);
        this.#imageLimits = imagePixelLimits(imageLimits);
    }

    /** Parses the profile HTML once; out-of-profile or malformed HTML throws. Images must be data URIs. */
    async render(document: HtmlPresentationDocument): Promise<Uint8Array> {
        // Before PDFKit decodes a PNG with transparency in JavaScript, or a GIF or WebP is decoded for it.
        checkImagePixels(document.html, this.#imageLimits);
        let result: Awaited<ReturnType<typeof renderHtml>>;

        try {
            result = await renderHtml(document.html, this.#faces);
        } catch (error) {
            // PDFKit writes no number at or past 1e21 (nor NaN) and throws its own message, which names no input: a
            // length in the HTML, or one it is multiplied into (an em thickness times the font size), is that large.
            if (error instanceof Error && error.message.startsWith(PDFKIT_UNSUPPORTED_NUMBER)) {
                throw new Error(`invalid PdfLapas HTML: a length, or one computed from it, is ${error.message.slice(PDFKIT_UNSUPPORTED_NUMBER.length)}, which a PDF cannot hold; keep the document's lengths near the slide's size`, { cause: error });
            }

            throw error;
        }

        this.#warnings = result.warnings;
        this.#missingFonts = result.missingFonts;
        if (result.warnings.length && this.#onUnsupported !== "warn") {
            throw new Error(`The PDFKit renderer cannot draw everything in this presentation (pass { onUnsupported: "warn" } to render anyway):\n- ${result.warnings.join("\n- ")}`);
        }

        reportMissingFonts(result.missingFonts, this.#onMissingFont);

        return result.pdf;
    }

    /** What the last render could not draw. */
    warnings(): string[] {
        return [...this.#warnings];
    }

    /** The font faces the last render asked for that none of the renderer's fonts provides. */
    missingFonts(): MissingFont[] {
        return [...this.#missingFonts];
    }
}

export async function createPdfkitRenderer(options: PdfkitRendererOptions): Promise<PdfkitRenderer> {
    return new PdfkitRenderer(options);
}

export async function createPdfkitRendererPool(_options: PdfkitRendererPoolOptions): Promise<PdfkitRendererPool> {
    throw new Error("createPdfkitRendererPool() starts worker threads and is only available under Node/Bun; in a browser use createPdfkitRenderer(), in a Web Worker to keep the page responsive");
}

export async function loadSystemFonts(_options?: SystemFontOptions): Promise<Uint8Array[]> {
    throw new Error("loadSystemFonts() reads the filesystem and is only available under Node/Bun; in a browser pass font bytes to createPdfkitRenderer({ fonts })");
}
