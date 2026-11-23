import { CSS_PX_PER_INCH, MILLIMETERS_PER_INCH, type NormalizedDecorationPiece } from "@pdflapas/model";
import type { DecorationRules, GenericFallback, RenderingTargetProfile, UnderlineShape } from "./index";

/** LibreOffice's hairline, one unit of its 1/100 mm model, in px: the line it strokes a wave with. */
const HAIRLINE_PX = CSS_PX_PER_INCH / (MILLIMETERS_PER_INCH * 100);
/** Its heavy wave's line, three of them (`wavyHeavy`). */
const HEAVY_HAIRLINES = 3;
/** Its double underline's and double strike's lines: each 0.64 T, the first centred T above the single line's centre, the second 2T below. */
const DOUBLE_LINES: readonly NormalizedDecorationPiece[] = Object.freeze([
    { kind: "line", top: -0.46, height: 0.64 },
    { kind: "line", top: 1.46, height: 0.64 },
]);
const line = (dashes?: readonly number[]): UnderlineShape => ({ pieces: [{ kind: "line", top: 0, height: 1, ...(dashes ? { dashes } : {}) }] });
/** A heavy type's line: twice as thick about the same centre, its dashes twice as long. */
const heavyLine = (dashes?: readonly number[]): UnderlineShape => ({ pieces: [{ kind: "line", top: -0.5, height: 2, ...(dashes ? { dashes: dashes.map(length => 2 * length) } : {}) }] });
const wave = (stroke: number): UnderlineShape => ({ pieces: [{ kind: "wave", centre: 0.5, amplitude: 2, halfPeriod: 4, stroke }] });
// Dash patterns, in thicknesses: dots one long, dashes 4 and long dashes 8, gaps 1 after a dot alone and 2 otherwise.
const DOTS = [1, 1];
const DASHES = [4, 2];
const LONG_DASHES = [8, 2];
const DOT_DASHES = [4, 2, 1, 2];
const DOT_DOT_DASHES = [4, 2, 1, 2, 1, 2];

/**
 * LibreOffice 26.2.6's underlines and strikes, read from its PDF export of a synthetic PptxGenJS 4.0.1 deck ("Lorem ipsum
 * dolor", each underline type, `sngStrike` and `dblStrike`, in Arial and Times New Roman at 12, 24 and 48pt, regular
 * and bold), where it strokes each as a line of a dash pattern or a wave: vcl's text line sizes from the face's
 * descent d, the same in bold. Underline and strike are d/4 thick (48pt Arial: 2.608pt, d/4 2.543), the underline's
 * centre d/2 below the baseline (5.187pt, 5.087) and the strike's (1 - d)/3 above it (12.614pt, 12.609), each within
 * 0.12pt at 48pt, its 1/100 mm rounding. Around that band (thickness T): a double underline or strike is two 0.64T
 * lines centred T above and 2T below the single one's centre (1.587pt at 3.6 and 8.391pt), heavy one 2T line on the
 * same centre (5.102pt), and each dash pattern is in thicknesses (DOTS, DASHES...), doubled with the heavy line; it
 * draws `dashLongHeavy` light (1.304pt, 8T and 2T at 24pt). `wavy` is a hairline wave about the centre, cubic
 * half-waves 4T long whose second control point is 2T off the line (24pt Arial: 5.216 and 2.608pt), `wavyHeavy` the
 * same three hairlines wide, `wavyDbl` two hairline waves 4/3 T off their lines, 8/3 T long, centred T/3 and T below
 * the centre (3.033 and 3.969pt at 24pt). A wave runs past the text to end its last half-wave. `words` draws no
 * underline. A dash pattern restarts with each run, and runs of different sizes keep their own bands.
 */
const LIBREOFFICE_DECORATION: DecorationRules = Object.freeze<DecorationRules>({
    bands: Object.freeze({ kind: "descent" as const, thickness: 0.25, underlineCentre: 0.5, strikeCentre: 1 / 3 }),
    underlines: Object.freeze({
        sng: line(),
        dbl: { pieces: DOUBLE_LINES },
        heavy: heavyLine(),
        dotted: line(DOTS),
        dottedHeavy: heavyLine(DOTS),
        dash: line(DASHES),
        dashHeavy: heavyLine(DASHES),
        dashLong: line(LONG_DASHES),
        dashLongHeavy: line(LONG_DASHES),
        dotDash: line(DOT_DASHES),
        dotDashHeavy: heavyLine(DOT_DASHES),
        dotDotDash: line(DOT_DOT_DASHES),
        dotDotDashHeavy: heavyLine(DOT_DOT_DASHES),
        wavy: wave(HAIRLINE_PX),
        wavyHeavy: wave(HEAVY_HAIRLINES * HAIRLINE_PX),
        wavyDbl: {
            pieces: [
                { kind: "wave", centre: 0.5 + 1 / 3, amplitude: 4 / 3, halfPeriod: 8 / 3, stroke: HAIRLINE_PX },
                { kind: "wave", centre: 1.5, amplitude: 4 / 3, halfPeriod: 8 / 3, stroke: HAIRLINE_PX },
            ],
        },
        words: undefined,
    }),
    doubleStrike: DOUBLE_LINES,
});

/**
 * Its fallback faces in the visual image (tests/visual/Dockerfile). The serif and monospace faces are those
 * fontconfig's 45-latin.conf files under them; every other face is sans-serif (49-sansserif.conf).
 */
const GENERIC_FALLBACK: GenericFallback = Object.freeze({
    serif: new Set([
        "bitstream vera serif",
        "cambria",
        "constantia",
        "dejavu serif",
        "elephant",
        "garamond",
        "georgia",
        "liberation serif",
        "luxi serif",
        "ms serif",
        "nimbus roman no9 l",
        "nimbus roman",
        "palatino linotype",
        "thorndale amt",
        "thorndale",
        "times new roman",
        "times",
    ]),
    monospace: new Set([
        "andale mono",
        "bitstream vera sans mono",
        "consolas",
        "courier new",
        "courier std",
        "courier",
        "cumberland amt",
        "cumberland",
        "dejavu sans mono",
        "fixedsys",
        "inconsolata",
        "liberation mono",
        "luxi mono",
        "nimbus mono l",
        "nimbus mono",
        "nimbus mono ps",
        "terminal",
    ]),
    emojiFace: "Noto Color Emoji",
    faces: Object.freeze({ "sans-serif": ["DejaVu Sans"], serif: ["DejaVu Serif", "DejaVu Sans"], monospace: ["DejaVu Sans Mono", "DejaVu Sans"] }),
    eastAsianFace: Object.freeze({ "sans-serif": "Noto Sans CJK", serif: "Noto Serif CJK", monospace: "Noto Sans Mono CJK" }),
    regionByLanguage: Object.freeze({ "zh-cn": "SC", "zh-hans": "SC", "zh-tw": "TC", "zh-hant-tw": "TC", zh: "HK", yue: "HK", ja: "JP", ko: "KR" }),
    otherRegion: "SC",
    sansSerifRegions: new Set(["HK"]),
});

/**
 * LibreOffice Impress, the visual reference (tests/visual): it kerns every run whatever its `kern`. A cell's
 * minimum height is its text height plus one unit of its 1/100 mm model, then the margins (svx table
 * Cell::getMinimumHeight, `GetTextHeight()+1`). A numbered paragraph PptxGenJS writes without a run (from a
 * run starting with "\n") restarts the numbering: three such items after an intro are all "I." in its PDF export.
 * Each line feed PptxGenJS keeps inside `<a:t>` (a run ending with "\n") starts a line that takes its paragraph
 * properties from the run that starts it (`lineFeedParagraph`). Its PPTX import draws a run with a baseline
 * attribute at DFLT_ESC_PROP, 58% of its size (oox TextCharacterProperties): measured on the visual runs slide,
 * 13pt "super" is 24 px wide against 42 px unscaled, "sub" 15 against 26 px (`escapementSizeRatio`).
 * Its synthetic styles, read from its PDF export of a regular-only face (Droid Sans Mono, 72pt): bold fills and
 * strokes the glyphs (`2 Tr`) with a line 1/30 of the size wide (`2.39996 w`); italic shears them
 * (`1 0 0.3462535606 1 x y Tm`).
 * It grows a line by what a raised or lowered run sticks out of it (`escapementGrowsLine`, #298): measured in its
 * PDF export (Arial 40pt, top-anchored, "H" + an escaped "x" + "H", then "H"), each run reaches the font-independent
 * line box at its drawn size `s`, `1.0 s` above and `0.2 s` below its baseline, shifted by its escapement `e`: the
 * line's ascent is the largest `1.0 s + e`, its descent the largest `0.2 s - e`. `baseline: 1000` (+20pt) puts the
 * first baseline 3.17pt lower (20 + 23.2 - 40 = 3.2), `baseline: -600` (-12pt) makes the line 8.62pt taller
 * (12 + 4.64 - 8 = 8.64), `subscript` 12.61pt, `baseline: 600` and `superscript` (+12pt) nothing.
 * It justifies Han text by widening every character too (`justifiesAtSpaces` off): in its PDF export of a synthetic
 * deck (SimSun 10pt, 3.33in boxes), a justified Han line with one space has 10.2pt Han steps and a 5.49pt space.
 * It draws no inner shadow on a picture (`drawsPictureInnerShadow`): in its PDF export a 2in picture with
 * `<a:innerShdw blurRad="101600" dist="76200" dir="2700000">` (75% black) is the bare picture, while the same
 * picture beside it with the outer shadow has one.
 * It draws no outline on a chart marker (`drawsChartMarkerOutline`): in its PDF export of a synthetic line chart with
 * 14pt circles, `lineDataSymbolLineColor: "000000"` and `lineDataSymbolLineSize: 3`, each marker is a 14pt disc in the
 * series color (56px across at 288 dpi), with no black ring and no growth.
 * Its underlines and strikes are LIBREOFFICE_DECORATION's (#489, #397).
 * It draws a character the run's face lacks from faces that depend on the run face's generic family and the run's
 * language (`genericFallback`, #289): an emoji from Noto Color Emoji, then DejaVu Sans, DejaVu Serif or DejaVu Sans
 * Mono (the last two then DejaVu Sans), then Noto Sans CJK, Noto Serif CJK or Noto Sans Mono CJK in the language's
 * region, Noto Sans CJK HK in every family. Measured in its 26.2 PDF export in the visual image (fonts named by each
 * page's font resources), "x<c>x" at 20pt for 55 characters (symbols, arrows, dingbats, emoji, Han, kana, Hangul,
 * Arabic, Hebrew, Greek, Cyrillic) in Arial, Calibri, Times New Roman, Courier New and Segoe UI, and Han, kana and
 * Hangul under 23 `lang` tags in Arial, Calibri, Times New Roman, Cambria, Georgia and Consolas: ⇒ ℕ ↉ ₹ come from
 * DejaVu Serif in Times New Roman and Cambria and from DejaVu Sans Mono in Courier New, ① ✓ ☆ ★, which neither has,
 * from DejaVu Sans; ❤ ✔ ⚠ ☀ ♠, which DejaVu Sans has, from Noto Color Emoji; 这 あ 한 from Noto Serif CJK SC in Times
 * New Roman, from Noto Sans CJK SC under en-US, fr-FR, zh-CN or zh-Hans, TC under zh-TW, HK under zh, zh-HK, zh-MO,
 * zh-SG or zh-Hant (Noto Sans CJK HK in Times New Roman and Consolas too), JP under ja and ja-JP, KR under ko, ko-KR and ko-KP. A face
 * it has no font for (Segoe UI, Georgia, Consolas) is drawn in the family's DejaVu face, which the chain then gives
 * Chromium too.
 * shortcut: Ⓐ comes from Noto Sans CJK HK whatever the family and language; under an HK region kana come from the JP
 * face and Hangul from the KR one, where the chain gives the HK face; ja-Latn takes SC, not JP. Each needs a chain per
 * character, worth it when a deck shows one.
 * shortcut: a face it has no font for keeps an emoji its DejaVu face has (❤ or 😀 in Segoe UI) in that face, where the
 * HTML lists Noto Color Emoji first; the HTML cannot say which faces are installed.
 * It draws U+2011 NON-BREAKING HYPHEN with the face's U+2010 glyph where the face lacks U+2011 (`nonBreakingHyphen`
 * "hyphen", HarfBuzz's normalization): in its PDF export Arial, Times New Roman and Courier New (Liberation, which
 * have U+2010 but not U+2011) draw it in the run's face, Calibri (Carlito, which has neither) in DejaVu Sans.
 * It draws every doughnut's hole half its outer diameter (`readsDoughnutHoleSize` off): in LibreOffice 26.2's PDF
 * export of PptxGenJS doughnuts with `<c:holeSize>` 30, 50 and 75 the hole is 0.496, 0.500 and 0.501 of it.
 * It draws a standard area chart's first series in front (`firstAreaSeriesInFront`, #533): in LibreOffice 26.2.6's PDF
 * export of PptxGenJS area charts with two series (Lorem [3, 5, 4, 7], Ipsum [5, 2, 6, 4]) and three (with Dolor
 * [4, 4, 4, 4]), Lorem fills everything below its line, Ipsum shows only above it and Dolor only above both, though
 * PptxGenJS writes them in that order (`<c:idx>` and `<c:order>` 0, 1, 2) in one `<c:areaChart>`.
 * A chart point whose `<c:dPt>` has no fill takes the automatic one, the theme's accent1, not its series fill
 * (`chartPointKeepsSeriesFill` off, #507): in LibreOffice 26.2.6's PDF export of single-series PptxGenJS charts with
 * `chartColors: ["FF0000", "00AA00", "0000FF"]`, every bar3D box (each `<c:dPt>` an outline in its color) and every
 * bar of a bar chart with `lineSize: 0` (each `<c:dPt>` `<a:ln><a:noFill/>`) is 4472C4 blue, while the legend key
 * stays FF0000; a bar3D chart with one chartColors entry, which gets no `<c:dPt>`, is red. The legend keeps the series
 * entry on those charts, and on a bar chart whose `<c:dPt>`s are filled in four chartColors (`chartLegendListsPoints`
 * off, #545, #552).
 * It draws a top-right legend in the chart's top corner beside the plot area, which keeps a right legend's band
 * (`topRightLegend`): in LibreOffice 26.2's PDF export of synthetic PptxGenJS 4.0.1 column charts (1.6 to 5in tall,
 * with and without a title, 1 to 5 entries, 10 to 20pt), the first key's center is 0.83 legend font sizes below the
 * chart's top edge whatever the height, the title or the entry count, and the plot area's top does not move (#536).
 * A bar3D value axis takes about as many steps as its labels' height fits along its drawn length
 * (`bar3DLabelHeightsPerStep` 1): in LibreOffice 26.2.6's PDF export of 3, 5, 4, 8 in 14pt Arial, 0..8 by 1 on a
 * 136.5pt axis, by 2 on a 95pt one and 0..10 by 5 on a 54pt one. A single series or a stacked chart gets no depth-axis
 * label (`bar3DNamesSingleSeries` off): its PDF export of the chartBar3DAxis fixture slide prints none (#541).
 * A bubble's value label prints its size (`bubbleValueLabelShowsSize`): in LibreOffice 26.2.6's PDF export of a
 * synthetic PptxGenJS 4.0.1 bubble chart with `showValue` (y 3, 5, 4, 7, sizes 1, 2, 3, 4) the labels read 1 to 4 (#477).
 * A pie whose `<c:dPt>`s have no `<c:dLblPos>` gets its labels at mid-radius, as `ctr` (`pieUnsetLabelPosition`,
 * #558): in LibreOffice 26.2.6's PDF export of a synthetic PptxGenJS 4.0.1 pie (4.4 x 4.1in, 40/25/20/15, 12pt Arial
 * values) the "40" label's center is 91px from the ring's center without dataLabelPosition and 167px with "bestFit".
 * It wraps an outEnd pie label to 80% of the room beside its slice's rim point, at most a fifth of the plot area's
 * width (`outsidePieLabelWidth`, PieChart::createTextLabelShape, whose room on the right is the plot area's width less
 * the point's x from the chart's edge): on the same pie with showLabel and showPercent, "Lorem" / "40%" on the right
 * becomes "Lor" / "em" / "40" / "%", so its lines are under 24pt (the width of "40%") and at least 17.3pt ("Lor").
 * A combo chart's legend lists its area series first (`comboLegendListsAreasFirst`, #538): in its PDF export of
 * synthetic bars A, B then areas R, S, and a line L then areas R, S, the legends read R, S, A, B and R, S, L. A combo
 * whose first type is bar keeps its categories between the ticks (`barComboKeepsCategoriesBetweenTicks`) though
 * PptxGenJS writes `<c:crossBetween val="midCat"/>`: Q1's label is centered half a category step (35.5pt) right of the
 * value axis, where the line and area combo puts it on the axis.
 */
export const LIBREOFFICE: RenderingTargetProfile = Object.freeze({
    readsKern: false,
    contentRowSlackMm: 0.01,
    addsParagraphSpacing: false,
    spacesFirstParagraph: true,
    spacedLineBaseline: undefined,
    fontLineBaseline: false,
    emptyNumberedParagraph: "restarts",
    lineFeedParagraph: "run",
    spacingWholePoints: false,
    percentSpacingEm: 1,
    escapementSizeRatio: 0.58,
    escapementGrowsLine: true,
    fullSizeMark: false,
    advanceGridPt: undefined,
    baselineGridPerInch: undefined,
    wrapWidthGridPt: undefined,
    joinsOpeningBracket: false,
    hangingMarks: {},
    breaksAroundHanPunctuation: false,
    justifiesAtSpaces: false,
    syntheticBoldStrokeEm: 1 / 30,
    syntheticItalicShear: 0.3462535606,
    fallbackFaces: [],
    eastAsianFallback: { byFace: {}, otherFaces: [] },
    emojiFace: undefined,
    // Not measured: U+FE0F keeps the renderer's emoji presentation, as before the flag.
    selectsFaceText: false,
    genericFallback: GENERIC_FALLBACK,
    nonBreakingHyphen: "hyphen",
    drawsPictureInnerShadow: false,
    drawsChartMarkerOutline: false,
    decoration: LIBREOFFICE_DECORATION,
    readsDoughnutHoleSize: false,
    firstAreaSeriesInFront: true,
    chartPointKeepsSeriesFill: false,
    chartLegendListsPoints: false,
    topRightLegend: { kind: "beside" as const },
    bar3DLabelHeightsPerStep: 1,
    // ChartView lays a 2D chart out twice and scales its value axes again on the second pass, with at most as many
    // steps as labels fit along the drawn axis (VCartesianAxis::estimateMaximumAutoMainIncrementCount, 2 to 10). Its
    // gridlines in its 26.2 PDF export of 1018 synthetic column, line and area charts (8 to 28pt Arial, plot 40 to 400pt
    // tall, data 3..7 to 120..870) fit the label height rounded to the nearest whole step with 37 misses, each within
    // 0.3em of a threshold: 3, 5, 4, 7 at 10pt runs 0..8 by 2 on a 73pt axis and by 1 on an 86.5pt one. On 180 horizontal
    // bar charts (10 and 18pt, plot 58 to 367pt wide) the step is the first scale's widest label width, floored:
    // 0.3..0.9 at 10pt ("0.1", 13.7pt) steps by 0.2 on a 128pt axis and by 0.1 on a 139pt one, 230..770 at 18pt ("900",
    // 30pt) by 500 on 126pt, 200 on 188pt and 100 on 297pt.
    // shortcut: 12pt labels step a little sooner than the rule and a horizontal axis whose labels LibreOffice turns
    // (under about 100pt wide at 10pt) steps otherwise; measure them when a deck needs them (#585).
    valueAxisStepRoom: { vertical: { label: "height" as const, perStep: 1, rounding: 0.5 }, horizontal: { label: "width" as const, perStep: 1, rounding: 0 }, maxIncrements: 10 },
    // chart2 ScaleAutomatism rounds the data's bounds out to the step, then moves one a step further where the data comes
    // within 1/21 of the range of it (PPTX_DEFAULTS.chart.automaticScale).
    automaticScale: { kind: "borderExpansion" as const },
    // Its PDF export draws a scatter and a bubble chart's y axis at x 3 for catAxisCrossesAt 3 (#591).
    bubbleCategoryCrossingShift: 0,
    bar3DNamesSingleSeries: false,
    bubbleValueLabelShowsSize: true,
    // Its PDF export of a bar and an area chart with showSerName and showValue prints "North; 7" (#477).
    dataLabelFieldSeparator: "; ",
    // A radar axis takes at most two steps: 26.2.6 draws 0..10 by 5 for 2..7 in 2.5, 4.5 and 7in high charts at 8, 12
    // and 14pt labels, 0..40 by 20 for 2..23, 0..100 by 50 for 20..70, 0..1 by 0.5 for 0.2..0.8, -10..10 by 10 for
    // -5..7, and 2..7 with valAxisMinVal 2 (its PDF export of synthetic PptxGenJS 4.0.1 radar charts, #561).
    radarScale: { kind: "bar3D" as const, maxIncrements: 2 },
    // Its PDF export of bar and line combos draws a line type's markers in the line's color, though PptxGenJS writes
    // their fill as `chartColors[index of the series in the chart]` (#604).
    readsComboMarkerFill: false,
    // Its PDF export of pie, radar and bar3D charts with a plotArea fill and border draws neither (#479).
    drawsPlotAreaOnEveryType: false,
    // bar3D draws its value axis on a wall, where 26.2.6 draws neither an axis line nor tick marks whatever
    // valAxisLine* and valAxis*TickMark say (#561).
    drawsBar3DValueAxisLine: false,
    // Its PDF export of a scatter or bubble chart with `showDataTable` draws a single malformed row ("Series1" over the
    // first x value), which the renderers do not draw (#592).
    drawsScatterDataTable: true,
    // Its x axis takes the scatter scale over the x values, not toward 0 (#548).
    scatterXAxisScale: { kind: "xValues" as const },
    pieUnsetLabelPosition: "ctr",
    comboLegendListsAreasFirst: true,
    barComboKeepsCategoriesBetweenTicks: true,
    outsidePieLabelWidth: { room: 0.8, plotWidth: 0.2 },
    // Its PDF export of a synthetic PptxGenJS 4.0.1 column chart with eight categories and `catAxisLabelFrequency: 2`
    // (`<c:tickLblSkip val="2"/>`) prints all eight labels (#566).
    readsCategoryLabelSkip: false,
    // Its PDF export of a bar3D chart with four series and `serAxisLabelFrequency: 2` names all four (#581).
    readsSeriesLabelSkip: false,
    // Its PDF export of column and line charts with `showDataTable` and `dataTableFormatCode` "0.00" or "$#,##0.0" prints
    // the table's values as in the data labels' format, 10 and 12, not 10.00 or $12.0 (#583).
    readsDataTableFormatCode: false,
    // Its PDF export of combo charts with a bar type and two line types (#594): `lineSmooth` on the first line type
    // smooths both lines, on the second one neither.
    comboLineSmoothFromFirstLineType: true,
    // Its PDF export of a bar and line combo whose line type has `secondaryCatAxis` draws catAxes[1]'s labels over the
    // primary ones, in their own color, size and order (#594, #603): `maxMin` reverses the labels and not the line
    // series; `high` moves them above the plot; its title stands above the plot area.
    secondaryCategoryAxis: { placement: "overlay" as const, seriesFollowAxis: false },
    // Its PDF export of column charts with `valAxisDisplayUnit: "thousands"` and `valAxisDisplayUnitLabel` (values
    // 3000 to 7000) and `"hundreds"` (300 to 700) numbers the axes 0..8000 and 0..800, with no unit name (#566).
    readsValueAxisDisplayUnits: false,
    // Its PDF export of column and line charts with `showDataTable` (each border and the keys on, or each off) draws the
    // same table: the category names, then each series' name and values, with no border and no key (#567).
    drawsDataTableBorders: false,
    // Its PDF export of a combo chart with two percentStacked bar series ([3, 5, 4, 7] and [5, 2, 6, 4]) and a line
    // ([6, 9, 7, 10]) stacks the bars to 8, 7, 10 and 11 on a 0 to 12 axis, as with "stacked" (#537).
    comboPercentBarStack: "stacked",
    // Its PDF export of 6 x 4.4in bestFit pies (12pt Arial) draws a 70/15/8/4/3 ring with percentages, all inside their
    // slices, 373px wide at 96 dpi and a 40/25/20/15 one 395px: VDiagram::reduceToMinimumSize tests the labels on a
    // small first ring (#572).
    bestFitLabelsShrinkPie: true,
});
