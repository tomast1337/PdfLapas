// The WASM visual decks: compare.ts scores them against LibreOffice through toPdf() and pins each
// PDF's bytes (pdf-sha256.json); tests/htmlBytes.test.ts pins their toHtml() bytes.
import type pptxgen from "pptxgenjs";
import type { PuppeteerGen } from "../../../packages/pptxgenjs/src/node";
import {
    populateBulletTypeFixture,
    populateChartAxisLabelFixture,
    populateChartBar3DAxisFixture,
    populateChartCategoryOptionFixture,
    populateChartDataTableFixture,
    populateChartPieBestFitFixture,
    populateChartBarColorFixture,
    populateChartBarOrderFixture,
    populateChartDataLabelFixture,
    populateChartCategoryAxisFixture,
    populateChartFixture,
    populateChartAreaFixture,
    populateChartComboFixture,
    populateChartComboAreaPieLabelFixture,
    populateChartAreaOrderFixture,
    populateChartLabelPositionFixture,
    populateChartLegendOrderFixture,
    populateChartMarkerFixture,
    populateChartOptionAxisFixture,
    populateChartOptionLabelFixture,
    populateChartPieFixture,
    populateChartPlotAreaFixture,
    populateChartPlotBandFixture,
    populateChartFixedBoundFixture,
    populateChartPointColorFixture,
    populateChartScatterAxesFixture,
    populateChartScatterFixture,
    populateChartStackedFixture,
    populateChartTopRightLegendFixture,
    populateChartValueAxisFixture,
    populateComponentFixture,
    populateFallbackFaceFixture,
    populateImageFixture,
    populatePlaceholderFixture,
    populateSlidePropertiesFixture,
    populateTableOptionsFixture,
    populateTextOptionsFixture,
    TEXT_OPTION_PAGES,
} from "../fixture";
import { populateShapeDeck, populateTableDeck, populateTextDeck, SHAPE_PAGES, TABLE_PAGES, TEXT_PAGES } from "./deck";
import { EFFECT_PAGES, populateEffectDeck } from "./effect-deck";
import { EMOJI_PAGES, populateEmojiDeck } from "./emoji-deck";
import { FORMAT_PAGES, populateGifDeck, populateWebpDeck } from "./format-deck";
import { IMAGE_PAGES, populateImageDeck } from "./image-deck";
import { LAYOUT_PAGES, populateLayoutDeck } from "./layout-deck";

export type Populate = <T extends pptxgen | PuppeteerGen>(presentation: T) => T;
export type SlideObject = ReturnType<PuppeteerGen["getModel"]>["slides"][number]["objects"][number];
const isText = (object: SlideObject) => object.kind === "text";

// emoji: Noto Color Emoji / Noto Sans CJK fallback (issue #14), installed in the pinned image.
// `crop` picks the objects that get a crop each: text boxes, for the shapes deck every shape
// (with or without text), so labels do not dilute the geometry scores, for the image format decks
// every picture, for the tables deck every table and for the charts deck every chart.
export const DECKS: Record<string, { populate: Populate; pages: number; crop: (object: SlideObject) => boolean }> = {
    text: { populate: populateTextDeck, pages: TEXT_PAGES, crop: isText },
    emoji: { populate: populateEmojiDeck, pages: EMOJI_PAGES, crop: isText },
    // layout: tab stops. LibreOffice drops a <a:tabLst> whose <a:pPr> carries marL/indent, which
    // PptxGenJS always writes, so its reference never shows explicit stops: those boxes count on the
    // page score only and the Rust test `tabs_end_at_their_stops` pins their positions.
    layout: { populate: populateLayoutDeck, pages: LAYOUT_PAGES, crop: object => object.kind === "text" && object.text.paragraphs.every(p => p.tabStops.length === 0) },
    shapes: { populate: populateShapeDeck, pages: SHAPE_PAGES, crop: object => object.kind === "shape" || (object.kind === "text" && object.shape !== undefined) },
    tables: { populate: populateTableDeck, pages: TABLE_PAGES, crop: object => object.kind === "table" },
    images: { populate: populateImageDeck, pages: IMAGE_PAGES, crop: object => object.kind === "image" },
    // svg: the shared fixture's two image slides, whose pictures are SVG.
    svg: {
        populate: presentation => {
            populateImageFixture(presentation);

            return presentation;
        },
        pages: 2,
        crop: object => object.kind === "image",
    },
    gif: { populate: populateGifDeck, pages: FORMAT_PAGES, crop: object => object.kind === "image" },
    webp: { populate: populateWebpDeck, pages: FORMAT_PAGES, crop: object => object.kind === "image" },
    // effects: every shadowed or glowing object; the slide titles sit above y = 0.5in (48px).
    effects: { populate: populateEffectDeck, pages: EFFECT_PAGES, crop: object => object.style.geometry.y > 48 },
    // charts: the shared fixture's chart slides and its axis-label slide. Pie and doughnut have the chartPie deck; radar and
    // scatter are left to tests/wasm.test.ts: LibreOffice draws them too far from ECharts for the page limit.
    charts: {
        populate: presentation => {
            populateChartFixture(presentation);
            populateChartAxisLabelFixture(presentation);

            return presentation;
        },
        pages: 6,
        crop: object => object.kind === "chart",
    },
    // chartLabels: the shared fixture's data-label slide (#445); every chart gets a crop.
    chartLabels: {
        populate: presentation => {
            populateChartDataLabelFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartAxis: the shared fixture's category-axis slide (#445); every chart gets a crop.
    chartAxis: {
        populate: presentation => {
            populateChartCategoryAxisFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartPlot: the shared fixture's plot-area and layout slide (#445); every chart gets a crop.
    chartPlot: {
        populate: presentation => {
            populateChartPlotAreaFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartMarker: the shared fixture's dash and dot marker slide (#484); every chart gets a crop.
    chartMarker: {
        populate: presentation => {
            populateChartMarkerFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartScatter: the shared fixture's scatter-line slide (#485); every chart gets a crop.
    chartScatter: {
        populate: presentation => {
            populateChartScatterFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartBarColors: the shared fixture's single-series bar color slide (#473); every chart gets a crop.
    chartBarColors: {
        populate: presentation => {
            populateChartBarColorFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartPie: the shared fixture's pie and doughnut slide (#468); every chart gets a crop.
    chartPie: {
        populate: presentation => {
            populateChartPieFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartBarOrder: the shared fixture's horizontal bar series-order slide (#472); every chart gets a crop.
    chartBarOrder: {
        populate: presentation => {
            populateChartBarOrderFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartLegendOrder: the shared fixture's stacked legend order slide (#508); every chart gets a crop.
    chartLegendOrder: {
        populate: presentation => {
            populateChartLegendOrderFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartStacked: the shared fixture's stacked bar width slide (#503); every chart gets a crop.
    chartStacked: {
        populate: presentation => {
            populateChartStackedFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartArea: the shared fixture's area fill slide (#520); every chart gets a crop.
    chartArea: {
        populate: presentation => {
            populateChartAreaFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartPlotBand: the shared fixture's plot-area band slide (#528); every chart gets a crop.
    chartPlotBand: {
        populate: presentation => {
            populateChartPlotBandFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartCombo: the shared fixture's combo chart grouping slide (#529); the chart gets a crop.
    chartCombo: {
        populate: presentation => {
            populateChartComboFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartAreaOrder: the shared fixture's standard area order slide (#533); every chart gets a crop.
    chartAreaOrder: {
        populate: presentation => {
            populateChartAreaOrderFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartPointColors: the shared fixture's point outlines slide (#507); every chart gets a crop.
    chartPointColors: {
        populate: presentation => {
            populateChartPointColorFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartScatterAxes: the shared fixture's scatter axis scale slide (#519); every chart gets a crop.
    chartScatterAxes: {
        populate: presentation => {
            populateChartScatterAxesFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartTopRightLegend: the shared fixture's vertical legend slide (#536); every chart gets a crop.
    chartTopRightLegend: {
        populate: presentation => {
            populateChartTopRightLegendFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartValueAxis: the shared fixture's value axis step slide (#547); every chart gets a crop.
    chartValueAxis: {
        populate: presentation => {
            populateChartValueAxisFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartLabelPositions: the shared fixture's pie rim and negative bar label slide (#478); every chart gets a crop.
    chartLabelPositions: {
        populate: presentation => {
            populateChartLabelPositionFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // bulletType: the shared fixture's non-number bullet type slide (#513, #531); every text box gets a crop.
    bulletType: {
        populate: presentation => {
            populateBulletTypeFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: isText,
    },
    // fallbackFaces: the shared fixture's fallback faces slides (#289, #403); every text box gets a crop.
    fallbackFaces: {
        populate: presentation => {
            populateFallbackFaceFixture(presentation);

            return presentation;
        },
        pages: 2,
        crop: isText,
    },
    // chartFixedBound: the shared fixture's fixed value-axis bound slide (#556); every chart gets a crop.
    chartFixedBound: {
        populate: presentation => {
            populateChartFixedBoundFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartBar3DAxis: the shared fixture's 3D value axis slide (#541); every chart gets a crop.
    chartBar3DAxis: {
        populate: presentation => {
            populateChartBar3DAxisFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartOptionLabels: the shared fixture's data labels on every chart type slide (#477); every chart gets a crop.
    chartOptionLabels: {
        populate: presentation => {
            populateChartOptionLabelFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartOptionAxes: the shared fixture's axis options on radar, bar3D and line charts slide (#561, #479); every
    // chart gets a crop.
    chartOptionAxes: {
        populate: presentation => {
            populateChartOptionAxisFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartComboAreaPieLabels: the shared fixture's combo area and pie label slides (#538, #558); every chart gets a crop.
    chartComboAreaPieLabels: {
        populate: presentation => {
            populateChartComboAreaPieLabelFixture(presentation);

            return presentation;
        },
        pages: 2,
        crop: object => object.kind === "chart",
    },
    // chartCategoryOptions: the shared fixture's category-axis and point options slide (#566); every chart gets a crop.
    chartCategoryOptions: {
        populate: presentation => {
            populateChartCategoryOptionFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartDataTable: the shared fixture's data tables, titles and depth axis slide (#567, #566); every chart gets a crop.
    chartDataTable: {
        populate: presentation => {
            populateChartDataTableFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // chartPieBestFit: the shared fixture's bestFit pie ring slide (#572); every chart gets a crop.
    chartPieBestFit: {
        populate: presentation => {
            populateChartPieBestFitFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "chart",
    },
    // components: the shared fixture's component slide (an HTML and a React progress bar, sized boxes;
    // the reference draws the same bars as PptxGenJS rectangles).
    components: {
        populate: presentation => {
            populateComponentFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "component",
    },
    // slides: the shared fixture's background picture, hidden slide and slide numbers (two printed pages);
    // the crops are the titles and the slide numbers, the picture is the page score.
    slides: {
        populate: presentation => {
            populateSlidePropertiesFixture(presentation);

            return presentation;
        },
        pages: 2,
        crop: isText,
    },
    // placeholders: the shared fixture's slide master placeholder slides (#465, #493); every filled placeholder gets a crop.
    placeholders: {
        populate: presentation => {
            populatePlaceholderFixture(presentation);

            return presentation;
        },
        pages: 2,
        crop: isText,
    },
    // options: the shared fixture's text-option slides; each case box (named after its parity crop) gets a crop, its label does not.
    options: {
        populate: presentation => {
            populateTextOptionsFixture(presentation);

            return presentation;
        },
        pages: TEXT_OPTION_PAGES,
        crop: object => object.style.objectName.startsWith("text-"),
    },
    // tableOptions: the shared fixture's table-option slide; every table gets a crop.
    tableOptions: {
        populate: presentation => {
            populateTableOptionsFixture(presentation);

            return presentation;
        },
        pages: 1,
        crop: object => object.kind === "table",
    },
};
