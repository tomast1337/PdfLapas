import type { DataUri, NormalizedPresentation, PageSizeInput } from "./types";

/**
 * The /Creator and /Producer every renderer writes to the PDF Info dictionary unless the document sets its own
 * (`PresentationMetadata.creator`, `producer`), as PowerPoint's Save as PDF names
 * itself in both. No version: a release would change every PDF's bytes (tests/visual/wasm/pdf-sha256.json), and the
 * crates have no version of their own. `crates/pdf-core/src/document_info.rs` keeps the same string as `PRODUCER`.
 */
export const PDF_PRODUCER = "PdfLapas";

export type HtmlPresentationDocument = {
    html: string;
    // Plain inches: a renderer also renders a document its caller built (#277).
    pageSize: PageSizeInput;
};

/** Turns the HTML document (`toHtml()`, `modelToHtml()`) into PDF bytes: every renderer reads only the HTML (#69). */
// eslint-disable-next-line @typescript-eslint/consistent-type-definitions -- a class implements this contract
export interface PresentationRenderer {
    render(document: HtmlPresentationDocument): Promise<Uint8Array>;
    /**
     * Optional: builds the HTML document from the model itself, by calling `modelHtml` from `request.htmlModule` (off
     * the caller's thread, #614), then renders it as render() would. toPdf() calls it instead of render(await toHtml())
     * when it is present and the deck has no components; the PDF must be the same bytes render() gives.
     */
    renderModelHtml?(request: ModelHtmlRequest): Promise<Uint8Array>;
}

/** What toPdf() hands PresentationRenderer.renderModelHtml(): plain data, cloned into another thread as it is. */
export type ModelHtmlRequest = {
    readonly model: NormalizedPresentation;
    /** Each path or URL image source of the model, resolved on the caller as toHtml() resolves it. */
    readonly images: ReadonlyMap<string, DataUri>;
    /** URL of a module whose `modelHtml(model, images)` export returns the toHtml() document for that model. */
    readonly htmlModule: string;
};

/**
 * Throws unless `document` carries the HTML string a renderer reads: `render(pptx.toHtml())` without `await` passes
 * the Promise, which would otherwise fail deep in the renderer (the WASM glue's "undefined is not an object").
 */
export function assertHtmlPresentationDocument(document: unknown): asserts document is HtmlPresentationDocument {
    if (typeof document === "object" && document !== null && "html" in document && typeof document.html === "string") {
        return;
    }

    const got = document instanceof Promise ? "a Promise: await toHtml() before passing its result" : `${document === null ? "null" : typeof document} without an html string`;

    throw new TypeError(`render() takes the { html, pageSize } document toHtml() resolves to, got ${got}`);
}
