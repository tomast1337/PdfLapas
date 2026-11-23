import type { NormalizedChart, NormalizedDecorationPiece } from "@pdflapas/model";
import { LIBREOFFICE } from "./libreoffice";
import { POWERPOINT } from "./powerpoint";

export { POWERPOINT_FACES, type PowerPointFace, type PowerPointFaceRole } from "./powerpoint";

/** The OOXML underline types (ST_TextUnderlineType) PptxGenJS writes as `<a:rPr u>`, but `none`. */
export const UNDERLINE_TYPES = Object.freeze(["sng", "dbl", "heavy", "dotted", "dottedHeavy", "dash", "dashHeavy", "dashLong", "dashLongHeavy", "dotDash", "dotDashHeavy", "dotDotDash", "dotDotDashHeavy", "wavy", "wavyHeavy", "wavyDbl", "words"] as const);
export type UnderlineType = (typeof UNDERLINE_TYPES)[number];

/** An underline type's shape: its pieces around the band. */
export type UnderlineShape = { readonly pieces: readonly NormalizedDecorationPiece[] };

/**
 * Where a target draws underlines and strikes, and their shapes around the band (docs/HTML_PROFILE.md). `bands`
 * "font": the face's `post` underline and OS/2 strikeout bands, the bold face's when bold, and the underlines of one
 * line share their bands averaged by width. "descent": both from the face's descent d (usWinDescent over unitsPerEm),
 * each `thickness` d tall, the underline's centre `underlineCentre` d below the baseline and the strike's
 * `strikeCentre` (1 - d) above it; each run keeps its own.
 */
export type DecorationRules = {
    readonly bands: { readonly kind: "font" } | { readonly kind: "descent"; readonly thickness: number; readonly underlineCentre: number; readonly strikeCentre: number };
    /** Each type's shape; undefined: the type draws no underline. */
    readonly underlines: Readonly<Record<UnderlineType, UnderlineShape | undefined>>;
    /** A double strike's (`dblStrike`) pieces; a single strike is one line filling the band. */
    readonly doubleStrike: readonly NormalizedDecorationPiece[];
};

/** A generic font family, as fontconfig sorts faces. */
export type GenericFamily = "sans-serif" | "serif" | "monospace";

/**
 * Fallback faces chosen by the run face's generic family and the run's language (normalize/fallback.ts
 * `fallbackChain`): the family's `faces`, then its East Asian face with the language's region appended.
 */
export type GenericFallback = {
    /** The lower-case faces of the serif family; a face in neither set is sans-serif. */
    readonly serif: ReadonlySet<string>;
    /** The lower-case faces of the monospace family. */
    readonly monospace: ReadonlySet<string>;
    /**
     * The face listed first for an emoji (Extended_Pictographic) whatever the run's face, which a renderer still takes
     * first where it has the emoji: the run is cut around each one (normalize/fallback.ts `fallbackSegments`).
     */
    readonly emojiFace: string;
    /** Each family's faces, first to last, before its East Asian face. */
    readonly faces: Readonly<Record<GenericFamily, readonly string[]>>;
    /** Each family's East Asian face, without its region ("Noto Sans CJK", to which "SC" is appended). */
    readonly eastAsianFace: Readonly<Record<GenericFamily, string>>;
    /** The region by lower-case language tag, else by its lower-case primary subtag, else `otherRegion`. */
    readonly regionByLanguage: Readonly<Record<string, string>>;
    readonly otherRegion: string;
    /** The regions whose East Asian face is the sans-serif family's whatever the run face's family. */
    readonly sansSerifRegions: ReadonlySet<string>;
};

/**
 * A rendering target: the application whose drawing of the PptxGenJS file the PDF matches, where applications differ.
 * Each quirk is a flag here; normalization reads the flags and never a target's name. Each target's values, and the
 * measurements behind them, live in its own file beside this one.
 */
export type RenderingTargetProfile = {
    /** Whether it honours `<a:rPr kern>` (PPTX_DEFAULTS.text.kernMinSizePt) or kerns every run. */
    readonly readsKern: boolean;
    /** What a row sized to its content (`<a:tr h="0">`) adds to its tallest cell's text and margins. */
    readonly contentRowSlackMm: number;
    /**
     * Whether a paragraph's space after and the next one's space before add up, or only the larger one counts (CSS
     * margins collapse).
     */
    readonly addsParagraphSpacing: boolean;
    /** Whether a text box's first paragraph keeps its space before. */
    readonly spacesFirstParagraph: boolean;
    /**
     * With explicit line spacing (`lineSpacing`, or a `lineSpacingMultiple` other than 1), each line's baseline below its
     * top as a fraction of the line height; undefined keeps `text.lineDescentEm` above its bottom.
     */
    readonly spacedLineBaseline: number | undefined;
    /**
     * Whether a single-spaced line's baseline comes from its fonts' FONT_WIN_METRICS, or sits `text.lineDescentEm` above
     * its bottom whatever the font. With it, spaced text also ends at its last baseline plus those fonts' descent, and a
     * line spacing as tall as the single line is single spacing.
     */
    readonly fontLineBaseline: boolean;
    /**
     * What a numbered paragraph PptxGenJS writes without a run (from a run starting with "\n") does to the numbering:
     * "restarts" it, or "skipped" keeps the current number without advancing it.
     */
    readonly emptyNumberedParagraph: "restarts" | "skipped";
    /**
     * Where each line feed PptxGenJS keeps inside `<a:t>` takes its paragraph properties from: "run", the run that starts
     * its line, or "paragraph", the `<a:p>`'s first `<a:pPr>`.
     */
    readonly lineFeedParagraph: "run" | "paragraph";
    /**
     * Whether `paraSpaceBefore`, `paraSpaceAfter` and point `lineSpacing` (`<a:spcPts>`) are rounded half-up to whole
     * points, each value on its own.
     */
    readonly spacingWholePoints: boolean;
    /**
     * The fraction of its font size a superscript or subscript run is drawn at; its baseline still moves by
     * `text.superscriptBaseline` or `text.subscriptBaseline` of the full size.
     */
    readonly escapementSizeRatio: number;
    /**
     * Whether a run raised or lowered off the baseline (super/subscript, custom `baseline`) grows its line where it
     * sticks out of it, or only its glyphs move and the line box stays where it would be without it.
     */
    readonly escapementGrowsLine: boolean;
    /**
     * Whether the paragraph mark takes its last run's size before that reduction and counts in its line's height, so a
     * paragraph ending in a large superscript, or the empty one after its last line feed, has a full line.
     */
    readonly fullSizeMark: boolean;
    /**
     * The grid each glyph's advance width (hmtx advance x size / unitsPerEm) is rounded to, to the nearest step, before
     * it lays text out; undefined lays it out with the natural advances.
     */
    readonly advanceGridPt: number | undefined;
    /**
     * The steps per inch of the grid each line's baseline is rounded to, to the nearest step; undefined keeps it where
     * the line box puts it.
     */
    readonly baselineGridPerInch: number | undefined;
    /**
     * The grid a text box's width inside its left and right insets is floored to before a line is tested against it;
     * undefined wraps at the width itself.
     */
    readonly wrapWidthGridPt: number | undefined;
    /**
     * Whether there is no line break before an opening bracket that follows a closing one or "%" (a U+2060 WORD JOINER
     * in the run text), where UAX #14 has one.
     */
    readonly joinsOpeningBracket: boolean;
    /**
     * By a run's primary language subtag, the marks that stay on their line past its end where they would start the
     * next one, unless another of them or a dash follows (normalize/eastAsianLineBreak.ts); a language without an
     * entry hangs none.
     */
    readonly hangingMarks: Readonly<Record<string, string>>;
    /**
     * Whether a run has line breaks around punctuation next to Han characters where UAX #14 has none (U+200B in the
     * run text) and none where it has one (U+2060): in a language with `hangingMarks`, before "“" and "‘" after Han,
     * between Han and an ASCII quote, and between two dashes but not before them; in any other, on either side of CJK
     * punctuation next to anything but white space, of full-width punctuation next to Han, and between Han and an
     * ASCII quote.
     */
    readonly breaksAroundHanPunctuation: boolean;
    /**
     * Whether a justified line takes its slack in its spaces only (`text-justify: inter-word`), every other character
     * at its own advance, and a line without a space stays at its start; otherwise the renderer's `auto`, which in
     * Chromium also widens each Han character.
     */
    readonly justifiesAtSpaces: boolean;
    /**
     * How a bold run is drawn in a face that has no bold style: the glyphs filled and stroked with a line this fraction
     * of the font size wide. Undefined leaves the browser's own synthesis (Chromium, measured in its PDF output of a
     * regular-only Liberation Sans at 9 to 72pt: a 1/32 em stroke, its glyph outlines 1/64 em wider on each side, and a
     * 0.25 shear), which the HTML does not write.
     */
    readonly syntheticBoldStrokeEm: number | undefined;
    /**
     * How an italic run is drawn in a face that has no italic style: the glyphs sheared to the right by this much per
     * unit of height above the baseline. Undefined leaves the browser's own synthesis (see `syntheticBoldStrokeEm`).
     */
    readonly syntheticItalicShear: number | undefined;
    /**
     * The faces a character the run's face lacks is drawn in, the first that has it, before the renderer's own
     * fallback (`NormalizedTextRun.fallbackFaces`); empty leaves it to the renderer.
     */
    readonly fallbackFaces: readonly string[];
    /**
     * The faces an East Asian character (`isEastAsian`) the run's face lacks is drawn in, after `fallbackFaces`: those
     * linked to the run's face (`byFace`, keyed by lower-case face), else `otherFaces`. Each such character keeps the
     * face of the one before it while that face has it, else takes the first of these that has it
     * (normalize/fallback.ts); empty leaves it to `fallbackFaces`.
     */
    readonly eastAsianFallback: { readonly byFace: Readonly<Record<string, readonly string[]>>; readonly otherFaces: readonly string[] };
    /** The face an emoji (Extended_Pictographic) the run's face lacks is drawn in, ahead of `fallbackFaces`; undefined: none. */
    readonly emojiFace: string | undefined;
    /** The faces after `fallbackFaces` by the run face's generic family and the run's language; undefined: none. */
    readonly genericFallback: GenericFallback | undefined;
    /**
     * How U+2011 NON-BREAKING HYPHEN is drawn: "hyphen-minus", with the run's face's U+002D glyph, "hyphen" as U+2010
     * from the first face that has it, both still unbreakable (`NormalizedTextRun.nonBreakingHyphen`); undefined, as
     * any other character.
     */
    readonly nonBreakingHyphen: "hyphen-minus" | "hyphen" | undefined;
    /**
     * Whether it draws a picture's inner shadow (`<a:innerShdw>` in its `<p:spPr>`, which PptxGenJS writes for
     * `addImage` with `shadow: { type: "inner" }`), or leaves the picture without it.
     */
    readonly drawsPictureInnerShadow: boolean;
    /**
     * Whether it draws the outline PptxGenJS writes around each line, radar and scatter marker (`<a:ln>` in
     * `<c:marker><c:spPr>`, `lineDataSymbolLineSize` in `lineDataSymbolLineColor`), or the marker without it.
     */
    readonly drawsChartMarkerOutline: boolean;
    /** How it draws underlines and strikes (`NormalizedTextDecoration.underlineShape`, `strikeShape`). */
    readonly decoration: DecorationRules;
    /**
     * Whether a doughnut's hole is `holeSize` percent of its outer diameter (`<c:holeSize>`, which PptxGenJS writes
     * from `holeSize`), or always PPTX_DEFAULTS.chart.doughnutHoleSizePct whatever it says.
     */
    readonly readsDoughnutHoleSize: boolean;
    /**
     * Whether a standard (unstacked) area chart, whose series overlap, draws its first series in front of the later
     * ones, or each series in front of the ones before it, in the order PptxGenJS writes them.
     */
    readonly firstAreaSeriesInFront: boolean;
    /**
     * Whether a chart point whose `<c:dPt>` has no fill (PptxGenJS's per-point outline on a single-series bar3D chart,
     * or `lineSize: 0` on a single-series bar chart with chartColors) keeps its series fill, or takes the automatic one
     * (`NormalizedChartSeries.automaticPointFill`).
     */
    readonly chartPointKeepsSeriesFill: boolean;
    /**
     * Whether the legend of a single-series bar or bar3D chart whose `<c:dPt>`s have no fill (PptxGenJS's per-point
     * outline, or `lineSize: 0`, with several chartColors) lists each point in its chartColors entry, or the series
     * (`NormalizedChart.legendListsPoints`).
     */
    readonly chartLegendListsUnfilledPoints: boolean;
    /**
     * Where it draws a top-right legend (`legendPos: "tr"`): in a band beside the plot area at the chart's top corner,
     * as a right legend's, or in a band above the plot area at the right edge (`NormalizedChart.topRightLegend`).
     */
    readonly topRightLegend: NormalizedChart["topRightLegend"];
    /**
     * How many value-label heights of drawn axis a bar3D value axis's automatic step needs: the step is the smallest
     * that leaves the axis at most its length over this many label heights in steps
     * (`NormalizedChartValueAxis.automaticScale.labelHeightsPerStep`).
     */
    readonly bar3DLabelHeightsPerStep: number;
    /**
     * How a 2D chart's automatic value axis fits its steps to the drawn axis (#585): at most `floor(length / (perStep ×
     * label) + rounding)` steps, between 2 and `maxIncrements`, where `label` is the labels' height or the widest
     * label's width, on a vertical axis and on a horizontal one (`NormalizedChartValueAxis.automaticScale.stepRoom`).
     */
    readonly valueAxisStepRoom: {
        readonly vertical: { readonly label: "height" | "width"; readonly perStep: number; readonly rounding: number };
        readonly horizontal: { readonly label: "height" | "width"; readonly perStep: number; readonly rounding: number };
        readonly maxIncrements: number;
    };
    /**
     * Whether a bar3D chart with one unstacked series prints its name at the depth axis, or no depth label
     * (`NormalizedChart.namesSingleDepthSlot`).
     */
    readonly bar3DNamesSingleSeries: boolean;
    /**
     * Whether a bubble chart's value label (`showValue`, `<c:showVal>`) prints the bubble's size, or its y value
     * (`NormalizedChartDataLabels.valueIsBubbleSize`).
     */
    readonly bubbleValueLabelShowsSize: boolean;
    /**
     * What joins a data label's series name and value, as PptxGenJS writes no `<c:separator>`
     * (`NormalizedChartDataLabels.fieldSeparator`).
     */
    readonly dataLabelFieldSeparator: string;
    /**
     * How many steps a radar chart's automatic value axis takes at most, with 0 included (`normalizeChart`'s radar
     * scale, the bar3D rounding over the data).
     */
    readonly radarMaxIncrements: number;
    /**
     * Where a pie's labels go when PptxGenJS writes no `<c:dLblPos>` (dataLabelPosition unset, or one it drops):
     * "bestFit" near the rim, or "ctr" at mid-radius.
     */
    readonly pieUnsetLabelPosition: "bestFit" | "ctr";
    /** Whether a combo chart's legend lists its area series first, or every series in call order (`NormalizedChart.legendListsAreasFirst`). */
    readonly comboLegendListsAreasFirst: boolean;
    /**
     * Whether a combo chart whose first type is bar keeps its categories between the tick marks, though PptxGenJS writes
     * `<c:crossBetween val="midCat"/>` on a combo with an area type, or puts them on the ticks (`NormalizedChart.categoriesOnTicks`).
     */
    readonly barComboKeepsCategoriesBetweenTicks: boolean;
    /** How wide a pie's outEnd label may be before it wraps (`NormalizedChart.outsidePieLabelWidth`); undefined: it never wraps. */
    readonly outsidePieLabelWidth: NormalizedChart["outsidePieLabelWidth"];
    /**
     * Whether a category axis prints a label every `catAxisLabelFrequency` categories (`<c:tickLblSkip>`), or every
     * label (`NormalizedChartCatAxis.labelSkip`).
     */
    readonly readsCategoryLabelSkip: boolean;
    /**
     * Whether a value axis divides its labels by `valAxisDisplayUnit` and prints the unit's name for
     * `valAxisDisplayUnitLabel` (`<c:dispUnits>`), or prints the values as they are (`NormalizedChartValueAxis.displayUnit`).
     */
    readonly readsValueAxisDisplayUnits: boolean;
    /**
     * Whether a chart's data table draws its borders and legend keys as `showDataTableHorzBorder`, `VertBorder`,
     * `Outline` and `Keys` say, or the text alone (`NormalizedChartDataTable`).
     */
    readonly drawsDataTableBorders: boolean;
    /**
     * How a combo chart's percent-stacked bars (`<c:grouping val="percentStacked"/>` on its bar element) stack
     * (`NormalizedChartSeries.stack`): by value, as `stacked` bars, or each as its share of its category, a fraction of
     * 1 on the axis it shares with the other types (`shareStacked`).
     */
    readonly comboPercentBarStack: "stacked" | "shareStacked";
    /**
     * Whether a bestFit pie label that does not fit in its slice of the first, small ring shrinks the ring, as an
     * outside label does (`NormalizedChart.bestFitLabelsShrinkPie`), or the ring keeps its full size.
     */
    readonly bestFitLabelsShrinkPie: boolean;
};

/** Every rendering target by the name the `target` option takes. A new target is one file here and one entry below. */
export const RENDERING_TARGETS = Object.freeze({
    libreoffice: LIBREOFFICE,
    powerpoint: POWERPOINT,
});
export type RenderingTarget = keyof typeof RENDERING_TARGETS;

/** `PuppeteerGen` without a `target`: the LibreOffice rendering that existing decks and the visual reference use. */
export const DEFAULT_RENDERING_TARGET: RenderingTarget = "libreoffice";
