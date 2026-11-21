import { type HtmlScan, scanHtml } from "./htmlScan";
import type { HtmlPresentationDocument } from "./renderer";
import { inches } from "./units";

/** The PdfLapas HTML profile version (docs/HTML_PROFILE.md) a document follows, in `<meta name="pdflapas-profile">`. */
export const HTML_PROFILE_VERSION = 2;
/** `PageSize.name` of a document whose `<meta name="pdflapas-page-size">` names none. */
export const CUSTOM_PAGE_SIZE_NAME = "Custom";

/** The prefix of the profile's `<meta>` names (`pdflapas-profile`, `pdflapas-page-size`, ...). */
const META_PREFIX = "pdflapas-";
const PROFILE_META = `<meta name="${META_PREFIX}profile" content="${HTML_PROFILE_VERSION}">`;
const PAGE_RULE = /@page\s*\{[^}]*?\bsize:\s*([^\s;}]+)in\s+([^\s;}]+)in\s*[;}]/;
const CSS_COMMENT = /\/\*[\s\S]*?\*\//g;

/** The `content` of the first `<meta name="<name>">`; none when there is no such `<meta>` or it has no `content`. */
function metaContent(scan: HtmlScan, name: string): { kind: "content"; content: string } | { kind: "none" } {
    const content = scan.metas.find(attributes => attributes.get("name") === name)?.get("content");

    return content === undefined ? { kind: "none" } : { kind: "content", content };
}

/**
 * The renderer input for a complete profile document: checks the version and reads the page size from the
 * `@page` rule, with the name `<meta name="pdflapas-page-size">` gives. Fills nothing: a producer completes its own
 * defaults first (PptxGenJS: `pptxHtmlDocument`). Throws when `html` is empty, has no
 * `<meta name="pdflapas-profile" content="2">`, is of another version or has no `@page { size: <w>in <h>in }`.
 */
export function htmlDocument(html: string): HtmlPresentationDocument {
    if (typeof html !== "string" || !html.trim()) {
        throw new Error(`invalid PdfLapas HTML: expected a non-empty HTML string, got ${typeof html === "string" ? "an empty string" : typeof html}`);
    }

    const scan = scanHtml(html);
    const prefix = profileMetaPrefix(name => {
        const meta = metaContent(scan, name);

        return meta.kind === "content" ? meta.content : undefined;
    });
    // The size of the first `@page` rule in the document's `<style>` elements, outside CSS comments.
    let match: RegExpExecArray | null = null;

    for (const css of scan.styles) {
        match = PAGE_RULE.exec(css.replace(CSS_COMMENT, ""));
        if (match) {
            break;
        }
    }

    const width = Number(match?.[1]);
    const height = Number(match?.[2]);

    if (!(width > 0 && height > 0 && Number.isFinite(width) && Number.isFinite(height))) {
        throw new Error("invalid PdfLapas HTML: the stylesheet's @page rule has no 'size: <width>in <height>in'; write the page size in inches, as <style>@page { size: 10in 5.625in; margin: 0 }</style> in <head>");
    }

    const pageSizeName = metaContent(scan, `${prefix}page-size`);

    return { html, pageSize: { width: inches(width), height: inches(height), name: pageSizeName.kind === "content" ? pageSizeName.content : CUSTOM_PAGE_SIZE_NAME } };
}

/**
 * The prefix of a profile document's `<meta>` names, `pdflapas-`, once the document declares the version readers accept.
 * `metaContent` reads the `content` of the document's `<meta name="<name>">`. Throws when the document declares no
 * profile version, or another one.
 */
// eslint-disable-next-line pdflapas/no-nullable-return -- phase 2 (#425)
export function profileMetaPrefix(metaContent: (name: string) => string | undefined): string {
    const version = metaContent(`${META_PREFIX}profile`);

    if (version === String(HTML_PROFILE_VERSION)) {
        return META_PREFIX;
    }

    if (version === undefined) {
        throw new Error(`invalid PdfLapas HTML: not a PdfLapas profile document: add ${PROFILE_META} to <head>. Puppeteer prints any HTML through renderer.render({ html, pageSize }).`);
    }

    throw new Error(`invalid PdfLapas HTML: profile version '${version}' in <meta name="${META_PREFIX}profile"> is not supported (this library reads ${PROFILE_META})`);
}
