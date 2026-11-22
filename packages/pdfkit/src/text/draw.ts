// Draws a laid-out text box: highlights, glyphs, decorations, links and bullets, on the profile's line boxes.
import { type Effect, sameEffect } from "#src/effect/index";
import { assertLinkUrl } from "@pdflapas/model";
import type { Face } from "#src/fonts";
import { captureFill, captureGlyphs, captureStroke, fill, glyphs, type Paint, stroke, useFace } from "#src/paint";
import { PDF_MITER_LIMIT } from "#src/path";
import { type BoxLayout, layoutText } from "./layout";
import type { Decoration, DecorationPiece, Run, TextBox } from "./read";
import { features } from "./shaper";

export type Frame = {
    readonly x: number;
    readonly y: number;
    readonly width: number;
    readonly height: number;
};

/** Lays the box's text out at `scale` in `width` px, the box's padding taken off. */
export function measureBox(paint: Paint, box: TextBox, width: number, scale = 1): BoxLayout {
    const missing = (run: Run) => paint.missingFace(run.face.family, paint.shaper.chain(run.face));

    return layoutText(box, width - box.padding[1] - box.padding[3], scale, paint.shaper, missing);
}

/** The content top of `layout` placed in `frame` by the box's anchor (`align-items`), the box's padding taken off. */
export function anchorTop(box: TextBox, layout: BoxLayout, frame: Frame): number {
    const [padTop, , padBottom] = box.padding;
    const slack = frame.height - padTop - padBottom - layout.height;
    const offset = box.anchor === "center" ? slack / 2 : box.anchor === "flex-end" ? slack : 0;

    return frame.y + padTop + offset;
}

/** Lays the box out in `frame`, shrinking it as its `data-text-fit` rule says; returns the layout and its content top. */
export function layoutBox(paint: Paint, box: TextBox, frame: Frame): { layout: BoxLayout; top: number; height: number } {
    const [padTop, padRight, padBottom, padLeft] = box.padding;
    const innerWidth = frame.width - padLeft - padRight;
    let innerHeight = frame.height - padTop - padBottom;
    let layout = measureBox(paint, box, frame.width);

    if (box.fit.kind === "shrink") {
        let scale = 1;

        for (let step = 0; step < box.fit.maxSteps; step++) {
            if (layout.height <= innerHeight + box.fit.tolerance && layout.width <= innerWidth + box.fit.tolerance) {
                break;
            }

            scale *= box.fit.step;
            layout = measureBox(paint, box, frame.width, scale);
        }
    }

    // `fit: resize` writes the height as min-height: the box grows to its text.
    if (box.fit.kind === "resize") {
        innerHeight = Math.max(innerHeight, layout.height);
    }

    const height = innerHeight + padTop + padBottom;

    return { layout, top: anchorTop(box, layout, { ...frame, height }), height };
}

/** A decoration band in px: its top `top` below the baseline, `thickness` tall. */
type Band = {
    readonly top: number;
    readonly thickness: number;
};

/** The font's metrics in px at `size`: hhea ascent and descent, post underline band, OS/2 strikeout band. */
function metrics(face: Face, size: number) {
    const font = face.font;
    const scale = size / font.unitsPerEm;
    const os2 = font["OS/2"] as { yStrikeoutPosition?: number; yStrikeoutSize?: number } | undefined; // `as`: the declared Os2Table omits these fields
    const underline: Band = { top: -font.underlinePosition * scale, thickness: font.underlineThickness * scale };
    const strike: Band = { top: -(os2?.yStrikeoutPosition ?? 0) * scale, thickness: (os2?.yStrikeoutSize ?? font.underlineThickness) * scale };

    return { ascent: font.ascent * scale, descent: -font.descent * scale, underline, strike };
}

/** The thinnest decoration band, px, as pdf-core draws it. */
const MIN_THICKNESS = 0.5;

/** A decoration's band in px at `size`: the HTML's em band, else the face's (`fallback`). */
function decorationBand(decoration: Decoration, fallback: Band, size: number): Band {
    return decoration.band ? { top: decoration.band.top * size, thickness: decoration.band.thickness * size } : fallback;
}

/** A run's part of a line a decoration is drawn under, from `start` to `end` px, on the baseline `y`. */
type Stretch = {
    readonly run: Run;
    readonly face: Face;
    readonly size: number;
    readonly y: number;
    readonly start: number;
    readonly end: number;
};

/**
 * Draws `decoration`'s pieces around `band` along `stretch` (docs/HTML_PROFILE.md "Underline and strike geometry"):
 * stripes and zigzags filled, waves stroked, lengths in the band's thickness (at least MIN_THICKNESS).
 */
function decorate(paint: Paint, decoration: Decoration, band: Band, stretch: Stretch): void {
    const t = Math.max(band.thickness, MIN_THICKNESS);
    const top = stretch.y + band.top;
    const { start, end } = stretch;
    const width = end - start;
    const filled: DecorationPiece[] = [];

    for (const piece of decoration.pieces) {
        if (piece.kind !== "wave") {
            filled.push(piece);
            continue;
        }

        // Cubic half-waves from the stretch's start, first control on their start and second at their middle, the
        // first up toward the baseline, as many as cover the stretch.
        const centre = top + piece.centre * t;
        const [half, amplitude] = [piece.half * t, piece.amplitude * t];
        const halves = Math.max(1, Math.ceil(width / half));
        const trace = (pen: PDFKit.PDFDocument) => {
            pen.moveTo(start, centre);
            for (let index = 0; index < halves; index++) {
                const from = start + half * index;

                pen.bezierCurveTo(from, centre, from + half / 2, index % 2 === 0 ? centre - amplitude : centre + amplitude, from + half, centre);
            }
        };

        paint.doc.save();
        trace(paint.doc);
        stroke(paint, decoration.color, piece.stroke);
        paint.doc.lineCap("butt").lineJoin("miter").miterLimit(PDF_MITER_LIMIT).undash().stroke();
        paint.doc.restore();
        captureStroke(paint, decoration.color.alpha, { width: piece.stroke, cap: "butt", join: "miter", miterLimit: PDF_MITER_LIMIT }, trace);
    }

    if (!filled.length) {
        return;
    }

    const trace = (pen: PDFKit.PDFDocument) => {
        for (const piece of filled) {
            if (piece.kind === "line") {
                const [pieceTop, height] = [top + piece.top * t, piece.height * t];

                if (!piece.dashes.length) {
                    pen.rect(start, pieceTop, width, height);
                    continue;
                }

                // Dashes on and off in turn from the stretch's start, the last one cut at its end.
                let x = start;

                for (let index = 0; x < end; index += 2) {
                    const on = (piece.dashes[index % piece.dashes.length] ?? 0) * t;

                    pen.rect(x, pieceTop, Math.min(on, end - x), height);
                    x += on + (piece.dashes[(index + 1) % piece.dashes.length] ?? 0) * t;
                }
            } else if (piece.kind === "zigzag") {
                // 45-degree slopes `line` px tall filling the piece's height, whole half-waves of `reach`, down first:
                // pdf-core's zigzag.
                const zigzagTop = top + piece.top * t;
                const reach = Math.max(piece.height * t - piece.line, piece.line);
                const halves = Math.max(1, Math.ceil(width / reach));
                const edge = (index: number) => (index % 2 === 0 ? zigzagTop : zigzagTop + reach);

                pen.moveTo(start, zigzagTop);
                for (let index = 0; index <= halves; index++) {
                    pen.lineTo(start + reach * index, edge(index) + piece.line);
                }

                for (let index = halves; index >= 0; index--) {
                    pen.lineTo(start + reach * index, edge(index));
                }

                pen.closePath();
            }
        }
    };

    trace(paint.doc);
    fill(paint, decoration.color);
    paint.doc.fill();
    captureFill(paint, decoration.color.alpha, trace);
}

/** pdf-core's cap on distinct glows per text box: each one is a mask of the whole box. */
const MAX_GLOWS_PER_BOX = 16;

/**
 * Draws `layout` with its content box's left edge at `left` and top at `top`, each distinct run glow under the
 * runs that glow with it (pdf-core's `draw_layout`).
 */
export function drawLayout(paint: Paint, layout: BoxLayout, left: number, top: number): void {
    const glows: Effect[] = [];

    for (const line of layout.lines) {
        for (const { run } of line.fragments) {
            if (run.glow && !glows.some(glow => run.glow && sameEffect(glow, run.glow))) {
                glows.push(run.glow);
            }
        }
    }

    if (glows.length > MAX_GLOWS_PER_BOX) {
        paint.report(`${glows.length - MAX_GLOWS_PER_BOX} glows past the first ${MAX_GLOWS_PER_BOX} in one text box are not drawn`);
    }

    const draw = () => {
        for (const line of layout.lines) {
            const baseline = top + line.baseline;

            for (const fragment of line.fragments) {
                const { run, face, size } = fragment;
                const doc = paint.doc;
                const x = left + fragment.x;
                const y = baseline + run.shift;
                const m = metrics(face, size);

                if (run.highlight) {
                    fill(paint, run.highlight);
                    const trace = (pen: PDFKit.PDFDocument) => pen.rect(x, y - m.ascent, fragment.width, m.ascent + m.descent);

                    trace(doc);
                    doc.fill();
                    captureFill(paint, run.highlight.alpha, trace);
                }

                useFace(paint, face, size);
                // The face lacks the run's bold or italic style: synthesize it as the HTML says (pdf-core's text/draw.rs).
                const embolden = run.face.bold && !face.bold;
                const oblique = run.face.italic && !face.italic;
                const options = { features: fragment.features, letterSpacing: run.letterSpacing };

                doc.save();
                if (oblique) {
                    doc.transform(1, 0, -run.synthesis.italicShear, 1, run.synthesis.italicShear * y, 0);
                }

                fill(paint, run.color);
                if (embolden) {
                    stroke(paint, run.color, size * run.synthesis.boldStroke);
                }

                glyphs(paint, fragment.text, x, y, { ...options, fill: true, stroke: embolden });
                // The silhouette is the filled glyphs, without the synthetic-bold or outline strokes, as pdf-core captures it.
                captureGlyphs(paint, face, size, fragment.text, x, y, { ...options, alpha: run.color.alpha, ...(run.glow ? { glow: run.glow } : {}) });
                if (run.outline) {
                    stroke(paint, run.outline.color, run.outline.width);
                    glyphs(paint, fragment.text, x, y, { ...options, fill: false, stroke: true });
                }

                doc.restore();
                const link = run.link;

                // A slide link carries href="#slide-N" too: data-slide decides, as in pdf-core.
                if (link?.slide !== undefined) {
                    doc.goTo(x, y - m.ascent, fragment.width, m.ascent + m.descent, `slide-${link.slide}`);
                } else if (link?.href !== undefined) {
                    assertLinkUrl(link.href, "a text run");
                    doc.link(x, y - m.ascent, fragment.width, m.ascent + m.descent, link.href);
                }
            }

            // The line's decoration stretches: each run's fragments merged into one (its consecutive fragments on the
            // line, so a pattern runs on across them).
            const runs: Stretch[] = [];

            for (const fragment of line.fragments) {
                const { run, face, size } = fragment;

                if (!run.underline && !run.strike) {
                    continue;
                }

                const [x, y] = [left + fragment.x, baseline + run.shift];
                const last = runs.at(-1);

                if (last?.run === run) {
                    runs.splice(-1, 1, { ...last, start: Math.min(last.start, x), end: Math.max(last.end, x + fragment.width) });
                } else {
                    runs.push({ run, face, size, y, start: x, end: x + fragment.width });
                }
            }

            // The band a `data-underline-line-band` underline takes: the bands of the line's runs carrying it,
            // averaged by their stretches' widths, as pdf-core's `line_underline_band`.
            let [bandWidth, bandTop, bandThickness] = [0, 0, 0];

            for (const stretch of runs) {
                const underline = stretch.run.underline;

                if (!underline?.lineBand) {
                    continue;
                }

                const band = decorationBand(underline, metrics(stretch.face, stretch.size).underline, stretch.size);
                const width = stretch.end - stretch.start;

                bandWidth += width;
                bandTop += width * band.top;
                bandThickness += width * band.thickness;
            }

            for (const stretch of runs) {
                const { run, face, size } = stretch;
                const m = metrics(face, size);
                const underline = run.underline;

                if (underline) {
                    const band = underline.lineBand && bandWidth > 0 ? { top: bandTop / bandWidth, thickness: bandThickness / bandWidth } : decorationBand(underline, m.underline, size);

                    if (stretch.end > stretch.start) {
                        decorate(paint, underline, band, stretch);
                    }
                }

                if (run.strike) {
                    decorate(paint, run.strike, decorationBand(run.strike, m.strike, size), stretch);
                }
            }
        }

        for (const { bullet, face, size, x, baseline } of layout.bullets) {
            useFace(paint, face, size);
            fill(paint, bullet.color);
            const options = { features: features(true, 0), letterSpacing: 0 };

            glyphs(paint, bullet.text, left + x, top + baseline, { ...options, fill: true, stroke: false });
            captureGlyphs(paint, face, size, bullet.text, left + x, top + baseline, { ...options, alpha: bullet.color.alpha });
        }
    };
    // Each glow captures only its own runs, so the glows nest around one drawing of the text.
    let nested: () => void = draw;

    for (const glow of glows.slice(0, MAX_GLOWS_PER_BOX)) {
        const inner = nested;

        nested = () => paint.effects.under(paint.doc, glow, true, paint.opacity, paint.report, inner);
    }

    nested();
}
