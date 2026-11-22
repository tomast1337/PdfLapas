import type { RenderingTargetProfile } from "./index";

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
 * Its underline is not the face's `post` table band (`underlineFromFont`): on the visual fixture's "Underline
 * styles" slide (Arial 20pt) its PDF export draws it thinner and lower, closer to Chromium's own underline than to
 * the band (text-underline-sng 0.0867 with the band, over the 0.08 crop target that Chromium's underline meets).
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
 * entry on those charts (`chartLegendListsUnfilledPoints` off, #545).
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
    nonBreakingHyphen: undefined,
    drawsPictureInnerShadow: false,
    drawsChartMarkerOutline: false,
    underlineFromFont: false,
    readsDoughnutHoleSize: false,
    firstAreaSeriesInFront: true,
    chartPointKeepsSeriesFill: false,
    chartLegendListsUnfilledPoints: false,
    topRightLegend: { kind: "beside" as const },
    bar3DLabelHeightsPerStep: 1,
    bar3DNamesSingleSeries: false,
    bubbleValueLabelShowsSize: true,
    // Its PDF export of a bar and an area chart with showSerName and showValue prints "North; 7" (#477).
    dataLabelFieldSeparator: "; ",
    // A radar axis takes at most two steps: 26.2.6 draws 0..10 by 5 for 2..7 in 2.5, 4.5 and 7in high charts at 8, 12
    // and 14pt labels, 0..40 by 20 for 2..23, 0..100 by 50 for 20..70, 0..1 by 0.5 for 0.2..0.8, -10..10 by 10 for
    // -5..7, and 2..7 with valAxisMinVal 2 (its PDF export of synthetic PptxGenJS 4.0.1 radar charts, #561).
    radarMaxIncrements: 2,
});
