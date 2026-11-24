import type PptxGenJS from "pptxgenjs";
import { PPTX_DEFAULTS } from "#src/defaults";
import { OOXML_PERCENT } from "#src/normalize/shape/path";

/**
 * What PptxGenJS 4.0.1 has written onto each `shadow` object of a deck, by object.
 * slideObjectToXml converts a text box's, shape's or picture's `shadow` to its OOXML values in place (`blur` and
 * `offset` to EMU, `angle` to 60000ths of a degree, `opacity` to 100000ths) and writes those: an object that shares the
 * options object with an earlier one converts the converted values again (#633). PdfLapas never writes onto the
 * caller's objects: each presentation keeps this record, as it does SharedRunOptions.
 */
export type SharedShadows = WeakMap<object, PptxGenJS.ShadowProps>;

/** The values slideObjectToXml writes into `<a:outerShdw>` (and onto the options object): EMU, 60000ths of a degree, 100000ths. */
export type ShadowXml = { readonly blurRad: number; readonly dist: number; readonly dir: number; readonly alpha: number };

export type WrittenShadow =
    /** Drawn from `shadow`, which holds the values the object writes in points, degrees and 0-1. */
    | { kind: "drawn"; shadow: PptxGenJS.ShadowProps }
    /** A later object sharing the options object writes `xml`, out of its schema range: none is drawn. */
    | { kind: "outOfRange"; xml: ShadowXml };

/** ST_PositiveFixedAngle: `dir` is below 21600000, one turn. */
const ANGLE_UNITS_PER_TURN = 360 * PPTX_DEFAULTS.units.angleUnitsPerDegree;

/** PptxGenJS's valToPts: a number of points as EMU, 0 for anything else. */
function pointsToEmu(points: unknown): number {
    return Math.round((Number(points) || 0) * PPTX_DEFAULTS.units.emuPerPoint);
}

/**
 * The shadow PptxGenJS writes for the next object holding `shadow`, recorded in `shared`. The first object writes the
 * caller's values; each later one the values the one before converted, converted again (slideObjectToXml, text and
 * picture branches, with `blur || 8`, `offset || 4`, `angle || 270`, `opacity || 0.75`). From the second object on
 * `dir` and `alpha` leave their schema range (`angle` 180 becomes 10800000 degrees, `opacity` 1 becomes 100000), and
 * LibreOffice 26.2 draws no shadow from such XML (a synthetic deck, three text boxes sharing one shadow object).
 * shortcut: PptxGenJS converts slide by slide when it writes, and PdfLapas counts in the order of the add calls; they
 * differ only when a deck adds an object to a slide after a later slide's object with the same shadow object.
 */
export function writeShadow(shared: SharedShadows, shadow: PptxGenJS.ShadowProps): WrittenShadow {
    // slideObjectToXml skips a "none" shadow, and converts nothing.
    if (shadow.type === "none") {
        return { kind: "drawn", shadow };
    }

    const earlier = shared.get(shadow);
    const read = earlier ?? shadow;
    const defaults = PPTX_DEFAULTS.shadow.objectWhenEnabled;
    const xml: ShadowXml = {
        blurRad: pointsToEmu(read.blur || defaults.blurPt),
        dist: pointsToEmu(read.offset || defaults.offsetPt),
        dir: Math.round((read.angle || defaults.angle) * PPTX_DEFAULTS.units.angleUnitsPerDegree),
        alpha: Math.round((read.opacity || defaults.opacity) * OOXML_PERCENT),
    };

    shared.set(shadow, { ...read, type: read.type || defaults.type, blur: xml.blurRad, offset: xml.dist, angle: xml.dir, opacity: xml.alpha, color: read.color || defaults.color });
    if (earlier && (xml.dir >= ANGLE_UNITS_PER_TURN || xml.alpha > OOXML_PERCENT)) {
        return { kind: "outOfRange", xml };
    }

    // What the object writes is `read` converted once, which normalizeShadow converts the same way.
    return { kind: "drawn", shadow: read };
}
