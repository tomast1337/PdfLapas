// The PDFKit internals the renderer needs past its public API: the current transform, the page's content stream and
// resources (to move drawn content under an effect, or into a transparency group), and objects built by hand (an
// effect's raw mask, a form XObject). pdfkit is pinned to an exact version in package.json, so these fields cannot
// move under the renderer.
import { apply, invert, type Matrix } from "./path";

type PdfReference = {
    write(data: Uint8Array): void;
    end(data?: Uint8Array): void;
};

type PdfkitInternals = {
    readonly _ctm: number[];
    readonly page: {
        readonly width: number;
        readonly height: number;
        readonly content: { readonly buffer: Uint8Array[] };
        readonly resources: PdfReference;
        readonly xobjects: Record<string, unknown>;
    };
    ref(data: Record<string, unknown>): PdfReference;
    addContent(data: string): void;
};

export function internals(doc: PDFKit.PDFDocument): PdfkitInternals {
    return doc as unknown as PdfkitInternals; // `as`: the fields above exist on every PDFKit 0.20 document
}

export function currentTransform(doc: PDFKit.PDFDocument): Matrix {
    const [a = 1, b = 0, c = 0, d = 1, e = 0, f = 0] = internals(doc)._ctm;

    return [a, b, c, d, e, f];
}

/**
 * Runs `draw` into a transparency group painted with `opacity`, so overlapping fills and strokes inside fade as one
 * (SVG's and CSS's `opacity`; pdf-core's `push_opacity`). The drawn content moves out of the page into a form
 * XObject that shares the page's resources.
 */
export function opacityGroup(doc: PDFKit.PDFDocument, opacity: number, draw: () => void): void {
    if (opacity >= 1) {
        draw();

        return;
    }

    const pdf = internals(doc);
    const content = pdf.page.content.buffer;
    const start = content.length;

    draw();
    const drawn = content.splice(start);
    const inverse = invert(currentTransform(doc));

    // Nothing drawn, or a degenerate transform that shows nothing.
    if (!drawn.length || inverse.kind === "singular") {
        return;
    }

    // The group's box: the whole page, in the user space the group is painted in.
    const toUser = inverse.matrix;
    const corners = [apply(toUser, 0, 0), apply(toUser, pdf.page.width, 0), apply(toUser, 0, pdf.page.height), apply(toUser, pdf.page.width, pdf.page.height)];
    const xs = corners.map(([x]) => x);
    const ys = corners.map(([, y]) => y);
    const form = pdf.ref({ Type: "XObject", Subtype: "Form", BBox: [Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)], Group: { Type: "Group", S: "Transparency" }, Resources: pdf.page.resources });

    for (const chunk of drawn) {
        form.write(chunk);
    }

    form.end();
    // Names only grow on a page, so the count names a new one.
    const name = `FG${Object.keys(pdf.page.xobjects).length + 1}`;

    pdf.page.xobjects[name] = form;
    doc.save();
    // A group XObject is painted with the fill alpha (ISO 32000-1 11.6.4.4).
    doc.fillOpacity(opacity);
    pdf.addContent(`/${name} Do`);
    doc.restore();
}

/** The part of a PDFKit image the ordering below wraps: a PNG keeps its png-js decoder in `image`, a JPEG has none. */
type OpenedImage = { image?: { decodePixels?: (use: (pixels: Uint8Array) => void) => void } };

/**
 * Makes `doc` write its pictures' objects in drawing order (#424). PDFKit embeds a PNG with an alpha channel, indexed
 * transparency or interlacing in the callback of an asynchronous inflate, and creates the soft mask's object and
 * writes the image there, so object numbers and offsets followed whichever inflate ended first. The inflates still run
 * concurrently; only their callbacks wait for those of the pictures drawn before. A failing callback fails the render.
 */
export function embedImagesInOrder(doc: PDFKit.PDFDocument): void {
    // `as`: openImage(src) is PDFKit 0.20's image loader (doc.image calls it), missing from @types/pdfkit 0.17.
    const pdf = doc as unknown as { openImage(src: unknown): OpenedImage };
    const open = pdf.openImage.bind(doc);
    let previous = Promise.resolve();

    pdf.openImage = src => {
        const opened = open(src);
        const png = opened.image;
        const decode = png?.decodePixels?.bind(png);

        if (png && decode) {
            png.decodePixels = use => {
                const pixels = new Promise<Uint8Array>(resolve => decode(resolve));
                const before = previous;

                previous = (async () => {
                    await before;
                    try {
                        use(await pixels);
                    } catch (error) {
                        doc.emit("error", error);
                    }
                })();
            };
        }

        return opened;
    };
}
