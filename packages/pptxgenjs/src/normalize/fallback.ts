import type { FontMetricsLookup, GlyphAdvancesResult } from "#src/defaults";
import { RENDERING_TARGETS, type RenderingTarget } from "#src/target/index";

/** A stretch of a run's text and the faces a character its face lacks is drawn in, first to last. */
export type FallbackSegment = {
    readonly text: string;
    readonly fallbackFaces: readonly string[];
};

/** The East Asian face the last character of a run was drawn in, which the next run of identical options keeps. */
export type FallbackHold = { face: string | undefined };

// Han, kana, bopomofo and Hangul, the CJK symbols, punctuation and compatibility blocks, and the full- and half-width
// forms: what PowerPoint draws from the East Asian faces it links to the run's face (eastAsianFallback).
const EAST_ASIAN = /[\p{Script=Han}\p{Script=Hiragana}\p{Script=Katakana}\p{Script=Bopomofo}\p{Script=Hangul}　-㏿︰-﹏＀-￯]/u;
// An emoji: a pictograph, or a flag of two regional indicators.
const PICTOGRAPHIC = /[\p{Extended_Pictographic}\p{Regional_Indicator}]/u;
// An emoji whose default presentation is emoji (😀 ⭐ ✅, a flag): the only ones PowerPoint takes from emojiFace.
const EMOJI_PRESENTATION = /^[\p{Emoji_Presentation}\p{Regional_Indicator}]/u;
// A variation selector or keycap mark stays with its emoji, as a renderer draws them.
const GRAPHEMES = new Intl.Segmenter(undefined, { granularity: "grapheme" });

/** Whether `character` is one PowerPoint draws from its East Asian faces when the run's face lacks it. */
export function isEastAsian(character: string): boolean {
    return EAST_ASIAN.test(character);
}

/** The target's East Asian faces for a run in `fontFace` (eastAsianFallback). */
function eastAsianFaces(fontFace: string, target: RenderingTarget): readonly string[] {
    const { byFace, otherFaces } = RENDERING_TARGETS[target].eastAsianFallback;
    const face = fontFace.toLowerCase();

    // Own keys only: a face named like an Object.prototype member is just another face.
    return (Object.hasOwn(byFace, face) ? byFace[face] : undefined) ?? otherFaces;
}

/**
 * The target's chain for a run in `fontFace` and `language`: `fallbackFaces`, then its generic family's faces
 * (`genericFallback`), then its East Asian faces, without `fontFace`.
 */
export function fallbackChain(fontFace: string, language: string, target: RenderingTarget): readonly string[] {
    const face = fontFace.toLowerCase();
    const { fallbackFaces, genericFallback } = RENDERING_TARGETS[target];
    const chain = [...fallbackFaces];

    if (genericFallback !== undefined) {
        const { serif, monospace, faces, eastAsianFace, regionByLanguage, otherRegion, sansSerifRegions } = genericFallback;
        const family = serif.has(face) ? "serif" : monospace.has(face) ? "monospace" : "sans-serif";
        const tag = language.toLowerCase();
        const primary = tag.split("-")[0] ?? tag;
        // Own keys only: a tag named like an Object.prototype member is just another language.
        const region = (Object.hasOwn(regionByLanguage, tag) ? regionByLanguage[tag] : undefined) ?? (Object.hasOwn(regionByLanguage, primary) ? regionByLanguage[primary] : undefined) ?? otherRegion;

        chain.push(...faces[family], `${eastAsianFace[sansSerifRegions.has(region) ? "sans-serif" : family]} ${region}`);
    }

    chain.push(...eastAsianFaces(fontFace, target));

    return chain.filter((value, index, all) => value.toLowerCase() !== face && all.indexOf(value) === index);
}

/**
 * `text` cut where PowerPoint draws a character the run's face lacks in another face than the chain's first that has
 * it, each stretch with its chain reordered so a renderer that takes the first face with the character does the same.
 * Measured against PowerPoint's PDF export of a synthetic deck (RENDERING_TARGETS.powerpoint.eastAsianFallback): an
 * East Asian character keeps the face the one before it was drawn in while that face has it; a space, a character of
 * the run's face or any other character ends that. `hold` carries that face in from the run before and out to the
 * next: the caller clears it between runs whose options differ, as PowerPoint starts over at a change of colour or
 * weight but carries the face into a run of identical options ("这"|"是" both Microsoft JhengHei). An emoji the run's
 * face lacks takes `emojiFace` when it has it and its default presentation is emoji (😀 ⭐ ✅, a flag); one whose
 * default is text (☺ ♥ ♠) follows the chain, and one the face has stays in the face: in the same export they are laid
 * out at the face's advances (☺ 1.021em in Arial and Times New Roman, 0.6em in Courier New) or, in Calibri, which lacks
 * ☺ ♥ ♠, at Segoe UI Symbol's (1, 0.685 and 0.685em), never at Segoe UI Emoji's 1.373em, and drawn black.
 * One stretch with the plain chain when every character takes the chain's first, or a face's characters are unknown.
 * shortcut: the held face goes on across a line break, where PowerPoint starts over (a wrapped "这是这是这是" starts
 * its second line's 是 in MS Gothic); the lines are the renderer's, so cut there when the profile states its breaks.
 * shortcut: "2⃣" (U+20E3 without U+FE0F) stays whole in the run's face, where PowerPoint draws U+20E3 in Segoe UI
 * Symbol; cut a combining mark the face lacks when a deck shows one.
 * shortcut: ☺ ♥ ♠ in Calibri take Arial, the chain's first with them, where PowerPoint takes Segoe UI Symbol; put it
 * first for text-default emoji when a deck needs the 0.02 to 0.24em.
 * A keycap with U+FE0F ("1️⃣") stays whole in the run's face; on a target with `selectsFaceText` its U+FE0F
 * becomes U+FE0E after this (normalize/text.ts), so a renderer draws the outline, not a colour emoji (#588).
 * With `genericFallback` (LibreOffice), no face's characters are looked up: `text` is cut around each emoji, whose
 * stretch lists `genericFallback.emojiFace` first.
 */
export function fallbackSegments(text: string, fontFace: string, bold: boolean, italic: boolean, language: string, target: RenderingTarget, fontMetrics: FontMetricsLookup, hold: FallbackHold = { face: undefined }): FallbackSegment[] {
    const chain = fallbackChain(fontFace, language, target);
    const { emojiFace, genericFallback } = RENDERING_TARGETS[target];
    const lookup = fontMetrics.glyphAdvances;
    const heldBefore = hold.face;

    // Nothing is held after a run whose faces are not looked up.
    hold.face = undefined;

    if (genericFallback !== undefined) {
        const emojiChain = [genericFallback.emojiFace, ...chain.filter(face => face !== genericFallback.emojiFace)];
        const stretches: { text: string; fallbackFaces: readonly string[] }[] = [];

        for (const { segment } of GRAPHEMES.segment(text)) {
            const fallbackFaces = PICTOGRAPHIC.test(segment) ? emojiChain : chain;
            const last = stretches.at(-1);

            if (last?.fallbackFaces === fallbackFaces) {
                last.text += segment;
            } else {
                stretches.push({ text: segment, fallbackFaces });
            }
        }

        // An empty text keeps its one run.
        return stretches.length ? stretches : [{ text, fallbackFaces: chain }];
    }

    if (lookup === undefined || !chain.length) {
        return [{ text, fallbackFaces: chain }];
    }

    const ownFont = lookup(fontFace, bold, italic);

    if (ownFont.kind === "unknown") {
        return [{ text, fallbackFaces: chain }];
    }

    const own = ownFont.advances;
    // Each face's characters, looked up once for the run.
    const characters = new Map<string, GlyphAdvancesResult>();
    const has = (face: string, codePoint: number) => {
        let font = characters.get(face);

        if (font === undefined) {
            font = lookup(face, bold, italic);
            characters.set(face, font);
        }

        return font.kind === "advances" && font.advances.advances.has(codePoint);
    };
    const eastAsian = eastAsianFaces(fontFace, target).filter(face => chain.includes(face));
    const segments: { text: string; first: string | undefined }[] = [];
    let held = heldBefore;

    for (const { segment } of GRAPHEMES.segment(text)) {
        const codePoint = segment.codePointAt(0) ?? 0;
        const missing = !own.advances.has(codePoint);
        const natural = missing ? chain.find(face => has(face, codePoint)) : undefined;
        const emoji = missing && emojiFace !== undefined && EMOJI_PRESENTATION.test(segment) && has(emojiFace, codePoint);
        const sticky = missing && !emoji && isEastAsian(segment);
        const first = emoji ? emojiFace : sticky ? ((held && has(held, codePoint) ? held : eastAsian.find(face => has(face, codePoint))) ?? natural) : natural;

        held = sticky ? first : undefined;
        // The chain's own first needs no stretch of its own.
        const key = first === natural ? undefined : first;
        const last = segments.at(-1);

        if (last && last.first === key) {
            last.text += segment;
        } else {
            segments.push({ text: segment, first: key });
        }
    }

    hold.face = held;

    if (segments.every(({ first }) => first === undefined)) {
        return [{ text, fallbackFaces: chain }];
    }

    return segments.map(({ text: piece, first }) => ({ text: piece, fallbackFaces: first === undefined ? chain : [first, ...chain.filter(face => face !== first)] }));
}
