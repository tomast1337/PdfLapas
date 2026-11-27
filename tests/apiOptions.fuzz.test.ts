// Property-based tests for the PptxGenJS-compatible API options (#661, second piece): whatever options a caller passes
// to addText(), addTable(), addShape(), addImage() or addChart(), the call and toHtml() either succeed or refuse with
// one of our documented errors. An error the JavaScript engine raises itself means a value reached code that did not
// check it, and the caller's objects are never written to; the same deck twice gives byte-identical HTML.
import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import type PptxGenJS from "pptxgenjs";
import { UnsupportedChartError } from "../packages/pptxgenjs/src/chart/errors";
import { PuppeteerGen } from "../packages/pptxgenjs/src/html";
import { InvalidModelError } from "../packages/pptxgenjs/src/modelSchema";
import { UNDERLINE_TYPES } from "../packages/pptxgenjs/src/target";
import { UnsupportedFeatureError } from "../packages/pptxgenjs/src/UnsupportedFeatureError";

// Every property fixes its seed, so a red build names the same counterexample on every machine until it is fixed.
const SEED = 661;
const RUNS = 1000;

// A 1x1 pixel PNG: the smallest valid image data URI, so addImage() exercises the image path without a decoder.
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

/**
 * Messages the JavaScript engine (or a built-in such as WeakSet) raises for a value it did not expect. Our own refusals
 * name the option and the fix (AGENTS.md "Fail loudly and helpfully"), so one of these escaping the API is an internal
 * failure, not a documented refusal.
 */
const INTERNAL_FAILURE = [
    /Cannot read propert/,
    /Cannot set propert/,
    /is not a function/,
    /is not iterable/,
    /Cannot convert undefined or null/,
    /Cannot destructure/,
    /Cannot use 'in' operator/,
    /Cannot assign to read only propert/,
    /Invalid array length/,
    /Invalid string length/,
    /Maximum call stack size exceeded/,
    /out of memory/i,
    /undefined is not an object/,
    /null is not an object/,
    /Cannot access .* before initialization/,
    /Reduce of empty array/,
];

/** Fails unless `error` is one of the documented errors the API throws for bad input; a JS-engine message is a bug. */
function expectDocumentedFailure(error: unknown): void {
    if (!(error instanceof Error)) {
        throw new Error(`the API threw a value that is not an Error: ${String(error)}`);
    }

    if (error instanceof UnsupportedFeatureError || error instanceof UnsupportedChartError || error instanceof InvalidModelError) {
        return;
    }

    for (const pattern of INTERNAL_FAILURE) {
        if (pattern.test(error.message)) {
            throw error;
        }
    }

    // Every other Error is a refusal our boundary validation raises: it names the option and how to fix it. The
    // RangeErrors from checkedSize() and the TypeErrors from the alignment and underline checks are deliberate.
    expect(error.message.length).toBeGreaterThan(0);
}

/** `value` frozen with everything it reaches; typed arrays cannot be frozen and are checked by value instead. */
function deepFreeze<T>(value: T): T {
    if (typeof value !== "object" || value === null || ArrayBuffer.isView(value) || Object.isFrozen(value)) {
        return value;
    }

    Object.freeze(value);
    for (const key of Reflect.ownKeys(value)) {
        // `as`: an own key of an object, read back from it.
        deepFreeze((value as Record<PropertyKey, unknown>)[key]);
    }

    return value;
}

// --- Values a caller can pass -----------------------------------------------------------------------------------

const EDGE_NUMBERS = [0, -1, 1, 0.5, 2, 100, 1000, -100, 1e6, -1e6, Number.MAX_SAFE_INTEGER, Number.EPSILON, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY];
// Structural counts are not magnitudes: a huge rowspan makes the grid loop that many times, so they stay small, with
// the non-finite values the API must refuse (the same bound a table would get from any caller).
const EDGE_COUNTS = [0, -1, 1, 2, 3, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY];

const magnitude = fc.oneof(fc.integer({ min: -10_000, max: 10_000 }), fc.double({ min: -1e9, max: 1e9, noNaN: true }), fc.constantFrom(...EDGE_NUMBERS));
const count = fc.oneof(fc.integer({ min: -2, max: 5 }), fc.constantFrom(...EDGE_COUNTS));
const wrongNumber = fc.constantFrom("x", "", "50%", null, true, false, [], {}, [1, 2]);
const numberOrWrong = fc.oneof({ weight: 7, arbitrary: magnitude }, { weight: 3, arbitrary: wrongNumber });

const coord = fc.oneof(
    fc.integer({ min: -1000, max: 2000 }),
    fc.double({ min: -1e9, max: 1e9, noNaN: true }),
    fc.constantFrom(...EDGE_NUMBERS),
    fc.constantFrom("50%", "0%", "100%", "-50%", "250%", "x", "", "50%", "NaN%"),
    // getSmartParseNumber reads anything else as 0, as PptxGenJS does.
    fc.constantFrom(null, true, false, [], {}),
);

const TEXT_LITERALS = ["", " ", "x", "Arial", 'Arial";}', "\n", "\t", "é", "漢字", "0".repeat(64), "FF0000", "50%", `line${String.fromCharCode(0x2028)}break`];
const text = fc.oneof(fc.string({ maxLength: 100 }), fc.constantFrom(...TEXT_LITERALS));

const HEX_OR_SCHEME = ["000000", "FFFFFF", "FF0000", "ff0000", "123456", "tx1", "tx2", "bg1", "bg2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "", "red", "#FF0000", "12345", "1234567", "red;background:url(x)"];
const color = fc.oneof(fc.constantFrom(...HEX_OR_SCHEME), fc.string({ maxLength: 20 }), numberOrWrong, fc.constantFrom({}, { color: "FF0000" }, { transparency: 50 }, { color: "x" }, { transparency: Number.NaN }, []));

const boolOrWrong = fc.oneof({ weight: 5, arbitrary: fc.boolean() }, { weight: 1, arbitrary: fc.constantFrom(0, 1, "true", "", null, {}, []) });
// Any option may also be absent, null (which PptxGenJS reads as absent) or of a type it does not declare.
const wrongType = fc.constantFrom(null, 5, -1, true, "x", {}, [], ["x"], [{}]);
const maybe = <T>(arbitrary: fc.Arbitrary<T>): fc.Arbitrary<unknown> => fc.oneof({ weight: 6, arbitrary }, { weight: 1, arbitrary: fc.constant(undefined) }, { weight: 1, arbitrary: wrongType });

const align = fc.oneof(fc.constantFrom("left", "center", "right", "justify", "x", "", "LEFT"), fc.constantFrom(0, "", false, null), wrongNumber);
const valign = fc.oneof(fc.constantFrom("top", "middle", "bottom", "x", "", "MIDDLE"), fc.constantFrom(0, "", false, null), wrongNumber);

const shadow: fc.Arbitrary<Record<string, unknown>> = fc.record({
    type: maybe(fc.constantFrom("outer", "inner", "none", "x", "")),
    opacity: maybe(numberOrWrong),
    blur: maybe(numberOrWrong),
    angle: maybe(numberOrWrong),
    offset: maybe(numberOrWrong),
    color: maybe(color),
    rotateWithShape: maybe(boolOrWrong),
});

const line = fc.oneof(
    fc.record({
        color: maybe(color),
        width: maybe(numberOrWrong),
        dashType: maybe(fc.constantFrom("solid", "dash", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot", "x")),
        beginArrowType: maybe(fc.constantFrom("none", "arrow", "diamond", "oval", "stealth", "triangle", "x")),
        endArrowType: maybe(fc.constantFrom("none", "arrow", "diamond", "oval", "stealth", "triangle", "x")),
        type: maybe(fc.constantFrom("none", "solid", "x")),
    }),
    fc.constantFrom("FF0000", 0, 1, true),
);

const fill = fc.oneof(fc.record({ color: maybe(color), transparency: maybe(numberOrWrong), type: maybe(fc.constantFrom("none", "solid", "x")), alpha: maybe(numberOrWrong) }), fc.constantFrom(0, 1, true, [], {}));

const bulletNumberType = fc.constantFrom("alphaLcParenBoth", "alphaLcParenR", "alphaLcPeriod", "alphaUcPeriod", "arabicParenR", "arabicPeriod", "romanLcParenR", "romanUcPeriod", "x", "");
const bullet = fc.oneof(
    fc.boolean(),
    fc.record({
        type: maybe(fc.oneof(fc.constantFrom("bullet", "number", "x", ""), numberOrWrong)),
        characterCode: maybe(fc.oneof(text, numberOrWrong)),
        indent: maybe(numberOrWrong),
        style: maybe(bulletNumberType),
        numberType: maybe(numberOrWrong),
        numberStartAt: maybe(numberOrWrong),
        startAt: maybe(numberOrWrong),
        code: maybe(fc.oneof(text, numberOrWrong)),
        marginPt: maybe(numberOrWrong),
    }),
);

const underline = fc.oneof(fc.boolean(), fc.constantFrom(...UNDERLINE_TYPES), fc.string({ maxLength: 8 }), fc.record({ style: maybe(fc.oneof(fc.constantFrom(...UNDERLINE_TYPES), fc.string({ maxLength: 8 }))), color: maybe(color) }));

const hyperlink = fc.record({
    url: maybe(fc.oneof(fc.constantFrom("https://example.com", "http://example.com", "mailto:a@example.com", "", "javascript:alert(1)"), fc.string({ maxLength: 30 }))),
    slide: maybe(numberOrWrong),
    tooltip: maybe(text),
});

const tabStop = fc.record({ position: maybe(numberOrWrong), alignment: maybe(fc.oneof(fc.constantFrom("l", "r", "ctr", "dec"), fc.string({ maxLength: 4 }), numberOrWrong)) });
const tabStops = fc.oneof({ weight: 6, arbitrary: fc.array(tabStop, { maxLength: 3 }) }, { weight: 1, arbitrary: fc.constant(undefined) });

const borderSide = fc.record({ type: maybe(fc.constantFrom("none", "solid", "dash", "x", "")), color: maybe(color), pt: maybe(numberOrWrong) });
const border = fc.oneof(borderSide, fc.tuple(borderSide, borderSide, borderSide, borderSide));

// --- Text options, shared by addText(), run options and table cells -------------------------------------------------

const textOptionFields = {
    x: maybe(coord),
    y: maybe(coord),
    w: maybe(coord),
    h: maybe(coord),
    // PptxGenJS ignores an option it does not know; so must PdfLapas.
    unknownOption: maybe(fc.jsonValue({ maxDepth: 1 })),
    fontSize: maybe(numberOrWrong),
    fontFace: maybe(text),
    color: maybe(color),
    align: maybe(align),
    valign: maybe(valign),
    bold: maybe(boolOrWrong),
    italic: maybe(boolOrWrong),
    underline: maybe(underline),
    bullet: maybe(bullet),
    charSpacing: maybe(numberOrWrong),
    margin: maybe(fc.oneof(magnitude, fc.array(numberOrWrong, { maxLength: 4 }))),
    fill: maybe(fill),
    line: maybe(line),
    shadow: maybe(shadow),
    glow: maybe(fc.record({ color: maybe(color), opacity: maybe(numberOrWrong), size: maybe(numberOrWrong) })),
    outline: maybe(fc.record({ color: maybe(color), size: maybe(numberOrWrong) })),
    transparency: maybe(numberOrWrong),
    rotate: maybe(numberOrWrong),
    wrap: maybe(boolOrWrong),
    breakLine: maybe(boolOrWrong),
    hyperlink: maybe(hyperlink),
    fit: maybe(fc.oneof(fc.constantFrom("none", "shrink", "resize", "x", ""), fc.string({ maxLength: 8 }))),
    isTextBox: maybe(boolOrWrong),
    shape: maybe(fc.oneof(fc.constantFrom("rect", "roundRect", "ellipse", "line", "triangle", "chevron", "x", ""), fc.string({ maxLength: 8 }))),
    rectRadius: maybe(numberOrWrong),
    indentLevel: maybe(numberOrWrong),
    lineSpacing: maybe(numberOrWrong),
    lineSpacingMultiple: maybe(numberOrWrong),
    paraSpaceAfter: maybe(numberOrWrong),
    paraSpaceBefore: maybe(numberOrWrong),
    strike: maybe(fc.oneof(fc.boolean(), fc.constantFrom("sngStrike", "dblStrike", "x", ""))),
    subscript: maybe(boolOrWrong),
    superscript: maybe(boolOrWrong),
    vert: maybe(fc.oneof(fc.constantFrom("horz", "vert", "vert270", "wordArtVert", "eaVert", "mongolianVert", "x", ""), fc.string({ maxLength: 8 }))),
    textDirection: maybe(fc.oneof(fc.constantFrom("horz", "vert", "vert270", "wordArtVert", "x", ""), fc.string({ maxLength: 8 }))),
    rtlMode: maybe(boolOrWrong),
    tabStops,
    lang: maybe(fc.oneof(fc.constantFrom("en-US", "zh-CN", "fr-CA", "x", ""), fc.string({ maxLength: 8 }))),
    highlight: maybe(color),
    baseline: maybe(numberOrWrong),
    objectName: maybe(text),
    placeholder: maybe(fc.string({ maxLength: 10 })),
};

const textOptions = fc.record(textOptionFields);
const cellOptions = fc.record({
    ...textOptionFields,
    border: maybe(border),
    colspan: maybe(count),
    rowspan: maybe(count),
});

const run = fc.record({ text: maybe(fc.string({ maxLength: 40 })), options: maybe(textOptions) });
const textInput = fc.oneof(fc.string({ maxLength: 200 }), fc.array(run, { maxLength: 3 }));

// --- Shape options ---------------------------------------------------------------------------------------------

const shapeName = fc.oneof(fc.constantFrom("rect", "roundRect", "ellipse", "line", "lineInv", "triangle", "star5", "chevron", "custGeom"), fc.string({ maxLength: 8 }));

const point = fc.oneof(
    fc.record({ x: numberOrWrong, y: numberOrWrong }),
    fc.record({ x: numberOrWrong, y: numberOrWrong, moveTo: boolOrWrong }),
    fc.record({ close: fc.constant(true) }),
    fc.record({ x: numberOrWrong, y: numberOrWrong, curve: fc.record({ type: fc.constant("cubic" as const), x1: numberOrWrong, y1: numberOrWrong, x2: numberOrWrong, y2: numberOrWrong }) }),
    fc.record({ x: numberOrWrong, y: numberOrWrong, curve: fc.record({ type: fc.constant("quadratic" as const), x1: numberOrWrong, y1: numberOrWrong }) }),
    fc.record({ x: numberOrWrong, y: numberOrWrong, curve: fc.record({ type: fc.constant("arc" as const), wR: numberOrWrong, hR: numberOrWrong, stAng: magnitude, swAng: magnitude }) }),
);

const shapeOptions = fc.record({
    x: maybe(coord),
    y: maybe(coord),
    w: maybe(coord),
    h: maybe(coord),
    unknownOption: maybe(fc.jsonValue({ maxDepth: 1 })),
    angleRange: maybe(fc.oneof(fc.tuple(numberOrWrong, numberOrWrong), fc.array(numberOrWrong, { maxLength: 2 }), numberOrWrong)),
    arcThicknessRatio: maybe(numberOrWrong),
    fill: maybe(fill),
    flipH: maybe(boolOrWrong),
    flipV: maybe(boolOrWrong),
    hyperlink: maybe(hyperlink),
    line: maybe(line),
    rectRadius: maybe(numberOrWrong),
    rotate: maybe(numberOrWrong),
    shadow: maybe(shadow),
    lineSize: maybe(numberOrWrong),
    lineDash: maybe(fc.oneof(fc.constantFrom("solid", "dash", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot", "x"), numberOrWrong)),
    lineHead: maybe(fc.oneof(fc.constantFrom("none", "arrow", "diamond", "oval", "stealth", "triangle", "x"), numberOrWrong)),
    lineTail: maybe(fc.oneof(fc.constantFrom("none", "arrow", "diamond", "oval", "stealth", "triangle", "x"), numberOrWrong)),
    shapeName: maybe(text),
    objectName: maybe(text),
});

const shapeInput = fc.record({ shapeName, options: shapeOptions, points: fc.array(point, { maxLength: 4 }) });

// --- Image options ---------------------------------------------------------------------------------------------

const imageOptions = fc.record({
    data: maybe(fc.oneof(fc.constant(PNG), fc.string({ maxLength: 40 }), numberOrWrong)),
    path: maybe(fc.oneof(fc.constantFrom("logo.png", "https://example.com/a.png", "missing", ""), fc.string({ maxLength: 20 }))),
    unknownOption: maybe(fc.jsonValue({ maxDepth: 1 })),
    altText: maybe(text),
    flipH: maybe(boolOrWrong),
    flipV: maybe(boolOrWrong),
    hyperlink: maybe(hyperlink),
    placeholder: maybe(fc.string({ maxLength: 10 })),
    rotate: maybe(numberOrWrong),
    rounding: maybe(boolOrWrong),
    shadow: maybe(shadow),
    sizing: maybe(fc.record({ type: maybe(fc.oneof(fc.constantFrom("contain", "cover", "crop", "x", ""), fc.string({ maxLength: 8 }))), w: maybe(coord), h: maybe(coord), x: maybe(coord), y: maybe(coord) })),
    transparency: maybe(numberOrWrong),
    objectName: maybe(text),
    x: maybe(coord),
    y: maybe(coord),
    w: maybe(coord),
    h: maybe(coord),
});

// --- Table options ---------------------------------------------------------------------------------------------

const tableOptions = fc.record({
    autoPage: maybe(boolOrWrong),
    unknownOption: maybe(fc.jsonValue({ maxDepth: 1 })),
    autoPageCharWeight: maybe(numberOrWrong),
    autoPageLineWeight: maybe(numberOrWrong),
    autoPageRepeatHeader: maybe(boolOrWrong),
    autoPageHeaderRows: maybe(count),
    autoPageSlideStartY: maybe(numberOrWrong),
    border: maybe(border),
    colW: maybe(fc.oneof(magnitude, fc.array(numberOrWrong, { maxLength: 4 }))),
    fill: maybe(fill),
    margin: maybe(fc.oneof(magnitude, fc.array(numberOrWrong, { maxLength: 4 }))),
    rowH: maybe(fc.oneof(magnitude, fc.array(numberOrWrong, { maxLength: 4 }))),
    verbose: maybe(boolOrWrong),
    newSlideStartY: maybe(numberOrWrong),
    fontFace: maybe(text),
    fontSize: maybe(numberOrWrong),
    color: maybe(color),
    align: maybe(align),
    valign: maybe(valign),
    bold: maybe(boolOrWrong),
    italic: maybe(boolOrWrong),
    underline: maybe(underline),
    bullet: maybe(bullet),
    charSpacing: maybe(numberOrWrong),
    highlight: maybe(color),
    indentLevel: maybe(numberOrWrong),
    lang: maybe(text),
    paraSpaceAfter: maybe(numberOrWrong),
    paraSpaceBefore: maybe(numberOrWrong),
    strike: maybe(fc.oneof(fc.boolean(), fc.constantFrom("sngStrike", "dblStrike", "x", ""))),
    subscript: maybe(boolOrWrong),
    superscript: maybe(boolOrWrong),
    transparency: maybe(numberOrWrong),
    textDirection: maybe(fc.oneof(fc.constantFrom("horz", "vert", "vert270", "wordArtVert", "x", ""), fc.string({ maxLength: 8 }))),
    vert: maybe(fc.oneof(fc.constantFrom("horz", "vert", "vert270", "wordArtVert", "x", ""), fc.string({ maxLength: 8 }))),
    wrap: maybe(boolOrWrong),
    breakLine: maybe(boolOrWrong),
    shadow: maybe(shadow),
    objectName: maybe(text),
    x: maybe(coord),
    y: maybe(coord),
    w: maybe(coord),
    h: maybe(coord),
});

const cell = fc.oneof(fc.string({ maxLength: 20 }), fc.integer(), fc.boolean(), fc.record({ text: maybe(fc.oneof(fc.string({ maxLength: 20 }), fc.array(run, { maxLength: 2 }))), options: maybe(cellOptions) }));
const tableRows = fc.array(fc.array(cell, { maxLength: 3 }), { minLength: 1, maxLength: 3 });

// --- Chart options ---------------------------------------------------------------------------------------------

const chartType = fc.constantFrom("bar", "line", "area", "pie", "doughnut", "radar", "scatter", "bubble", "bar3D", "unknownType");
const formatCode = fc.oneof(fc.constantFrom("#%", "0.00%", "$0.00", "0", "", "x"), fc.string({ maxLength: 8 }));
const gridLine = fc.record({ cap: maybe(fc.constantFrom("flat", "round", "square", "x")), color: maybe(color), size: maybe(numberOrWrong), style: maybe(fc.constantFrom("solid", "dash", "dot", "none", "x")) });

const valueAxis = fc.record({
    showValAxisTitle: maybe(boolOrWrong),
    valAxisTitle: maybe(text),
    valAxisMinVal: maybe(magnitude),
    valAxisMaxVal: maybe(magnitude),
    valAxisMajorUnit: maybe(magnitude),
    valAxisLogScaleBase: maybe(magnitude),
    valAxisOrientation: maybe(fc.constantFrom("minMax", "maxMin", "x")),
    valAxisDisplayUnit: maybe(fc.oneof(fc.constantFrom("billions", "millions", "thousands", "trillions", "x"), fc.string({ maxLength: 8 }))),
    secondaryValAxis: maybe(boolOrWrong),
    valAxisHidden: maybe(boolOrWrong),
    valAxisLabelFontSize: maybe(numberOrWrong),
    valAxisLabelColor: maybe(color),
    valAxisLineColor: maybe(color),
    valAxisLineSize: maybe(numberOrWrong),
    valAxisLabelFormatCode: maybe(formatCode),
});
const categoryAxis = fc.record({
    showCatAxisTitle: maybe(boolOrWrong),
    catAxisTitle: maybe(text),
    catAxisMinVal: maybe(magnitude),
    catAxisMaxVal: maybe(magnitude),
    catAxisMajorUnit: maybe(magnitude),
    catAxisOrientation: maybe(fc.constantFrom("minMax", "maxMin", "x")),
    catAxisLabelFrequency: maybe(fc.oneof(magnitude, fc.string({ maxLength: 4 }))),
    secondaryCatAxis: maybe(boolOrWrong),
    catAxisHidden: maybe(boolOrWrong),
    catAxisLabelFontSize: maybe(numberOrWrong),
    catAxisLabelColor: maybe(color),
    catAxisLineColor: maybe(color),
    catAxisLabelFormatCode: maybe(formatCode),
});

const chartOptionFields = {
    x: maybe(coord),
    y: maybe(coord),
    w: maybe(coord),
    h: maybe(coord),
    objectName: maybe(text),
    unknownOption: maybe(fc.jsonValue({ maxDepth: 1 })),
    chartColors: maybe(fc.array(fc.oneof(fc.constantFrom(...HEX_OR_SCHEME), fc.string({ maxLength: 8 })), { maxLength: 4 })),
    chartColorsOpacity: maybe(numberOrWrong),
    showLegend: maybe(boolOrWrong),
    legendPos: maybe(fc.oneof(fc.constantFrom("b", "l", "r", "t", "tr", "x", ""), fc.string({ maxLength: 4 }))),
    showValue: maybe(boolOrWrong),
    showTitle: maybe(boolOrWrong),
    title: maybe(text),
    titleColor: maybe(color),
    titleFontSize: maybe(numberOrWrong),
    titleBold: maybe(boolOrWrong),
    dataLabelPosition: maybe(fc.oneof(fc.constantFrom("b", "bestFit", "ctr", "l", "r", "t", "inEnd", "outEnd", "x"), fc.string({ maxLength: 8 }))),
    dataLabelFontSize: maybe(numberOrWrong),
    dataLabelColor: maybe(color),
    dataLabelFormatCode: maybe(formatCode),
    barGrouping: maybe(fc.oneof(fc.constantFrom("clustered", "stacked", "percentStacked", "standard", "x"), fc.string({ maxLength: 8 }))),
    barGapWidthPct: maybe(numberOrWrong),
    barOverlapPct: maybe(numberOrWrong),
    holeSize: maybe(numberOrWrong),
    firstSliceAng: maybe(numberOrWrong),
    radarStyle: maybe(fc.oneof(fc.constantFrom("standard", "marker", "filled", "x"), fc.string({ maxLength: 8 }))),
    lineSize: maybe(numberOrWrong),
    lineDataSymbolSize: maybe(numberOrWrong),
    lineDataSymbol: maybe(fc.oneof(fc.constantFrom("circle", "dash", "diamond", "dot", "none", "square", "triangle", "x"), fc.string({ maxLength: 8 }))),
    valAxisMinVal: maybe(magnitude),
    valAxisMaxVal: maybe(magnitude),
    valAxisMajorUnit: maybe(magnitude),
    valAxisLogScaleBase: maybe(magnitude),
    valAxes: maybe(fc.array(valueAxis, { maxLength: 2 })),
    catAxes: maybe(fc.array(categoryAxis, { maxLength: 2 })),
    catAxisMinVal: maybe(magnitude),
    catAxisMaxVal: maybe(magnitude),
    catAxisMajorUnit: maybe(magnitude),
    catAxisLabelFrequency: maybe(fc.oneof(magnitude, fc.string({ maxLength: 4 }))),
    catAxisLabelRotate: maybe(numberOrWrong),
    catAxisLabelPos: maybe(fc.oneof(fc.constantFrom("none", "low", "high", "nextTo", "x"), fc.string({ maxLength: 8 }))),
    catAxisCrossesAt: maybe(fc.oneof(magnitude, fc.constantFrom("autoZero", "min", "max", "x"))),
    valAxisCrossesAt: maybe(fc.oneof(magnitude, fc.constantFrom("autoZero", "min", "max", "x"))),
    catLabelFormatCode: maybe(formatCode),
    valLabelFormatCode: maybe(formatCode),
    dataTableFormatCode: maybe(formatCode),
    valAxisLabelFormatCode: maybe(formatCode),
    catGridLine: maybe(gridLine),
    valGridLine: maybe(gridLine),
    dataBorder: maybe(borderSide),
    fill: maybe(fill),
    shadow: maybe(shadow),
    layout: maybe(fc.record({ x: maybe(magnitude), y: maybe(magnitude), w: maybe(magnitude), h: maybe(magnitude) })),
    showDataTable: maybe(boolOrWrong),
    v3DPerspective: maybe(magnitude),
    v3DRotX: maybe(magnitude),
    v3DRotY: maybe(magnitude),
    displayBlanksAs: maybe(fc.oneof(fc.constantFrom("gap", "zero", "span", "x"), fc.string({ maxLength: 8 }))),
    titlePos: maybe(fc.record({ x: magnitude, y: magnitude })),
};

const chartOptions = fc.record(chartOptionFields);
const chartLabel = fc.oneof(fc.string({ maxLength: 20 }), numberOrWrong);
const chartSeries = fc.record({
    name: maybe(fc.oneof(fc.string({ maxLength: 20 }), numberOrWrong)),
    unknownSeriesOption: maybe(fc.jsonValue({ maxDepth: 1 })),
    labels: maybe(fc.oneof(fc.array(chartLabel, { maxLength: 4 }), fc.array(fc.array(chartLabel, { maxLength: 3 }), { maxLength: 3 }))),
    values: maybe(fc.array(fc.oneof(magnitude, wrongNumber), { maxLength: 6 })),
    sizes: maybe(fc.array(fc.oneof(magnitude, wrongNumber), { maxLength: 6 })),
});
const chartData = fc.array(chartSeries, { maxLength: 3 });
const chartInput = fc.record({ chartType, data: chartData, options: chartOptions });

// --- One call of the API -----------------------------------------------------------------------------------------

type ApiCall =
    | { readonly method: "text"; readonly text: unknown; readonly options: Record<string, unknown> }
    | { readonly method: "table"; readonly rows: unknown; readonly options: Record<string, unknown> }
    | { readonly method: "shape"; readonly shapeName: string; readonly options: Record<string, unknown>; readonly points: unknown }
    | { readonly method: "image"; readonly options: Record<string, unknown> }
    | { readonly method: "chart"; readonly chartType: string; readonly data: unknown; readonly options: Record<string, unknown> };

const textCall: fc.Arbitrary<ApiCall> = fc.record({ text: textInput, options: textOptions }).map(({ text, options }) => ({ method: "text", text, options }));
const tableCall: fc.Arbitrary<ApiCall> = fc.record({ rows: tableRows, options: tableOptions }).map(({ rows, options }) => ({ method: "table", rows, options }));
const shapeCall: fc.Arbitrary<ApiCall> = shapeInput.map(({ shapeName, options, points }) => ({ method: "shape", shapeName, options, points }));
const imageCall: fc.Arbitrary<ApiCall> = imageOptions.map(options => ({ method: "image", options }));
const chartCall: fc.Arbitrary<ApiCall> = chartInput.map(({ chartType, data, options }) => ({ method: "chart", chartType, data, options }));

const call = fc.oneof(textCall, tableCall, shapeCall, imageCall, chartCall);

// --- Building the deck ------------------------------------------------------------------------------------------

/** The deck `call` builds on one fresh presentation, which every property builds again from a copy of `call`. */
function buildDeck(call: ApiCall): PuppeteerGen {
    const pres = new PuppeteerGen();
    const slide = pres.addSlide();

    if (call.method === "text") {
        // `as unknown as`: the generated values are the plain JSON a JavaScript caller passes; the API checks them.
        slide.addText(call.text as string | PptxGenJS.TextProps[], call.options as unknown as PptxGenJS.TextPropsOptions);
    } else if (call.method === "table") {
        slide.addTable(call.rows as PptxGenJS.TableRow[], call.options as unknown as PptxGenJS.TableProps);
    } else if (call.method === "shape") {
        const options = call.shapeName === "custGeom" ? { ...call.options, points: call.points } : call.options;

        slide.addShape(call.shapeName as PptxGenJS.SHAPE_NAME, options as unknown as PptxGenJS.ShapeProps);
    } else if (call.method === "image") {
        slide.addImage({ data: PNG, ...call.options } as unknown as PptxGenJS.ImageProps);
    } else {
        slide.addChart(call.chartType as PptxGenJS.CHART_NAME, call.data as PptxGenJS.OptsChartData[], call.options as unknown as PptxGenJS.IChartOpts);
    }

    return pres;
}

/** The call and toHtml(), so a property sees both the API's refusal and the rendering's. */
async function runCall(call: ApiCall): Promise<void> {
    await buildDeck(call).toHtml();
}

/** What the deck's HTML is, or the documented refusal it raises; both builds of a deterministic call agree. */
async function outcome(call: ApiCall): Promise<{ readonly kind: "html"; readonly html: string } | { readonly kind: "error"; readonly name: string; readonly message: string }> {
    try {
        return { kind: "html", html: (await buildDeck(structuredClone(call)).toHtml()).html };
    } catch (error) {
        expectDocumentedFailure(error);

        // `as`: expectDocumentedFailure() has just established that this is an Error.
        const failure = error as Error;

        return { kind: "error", name: failure.name, message: failure.message };
    }
}

describe("PptxGenJS API options fuzz (#661)", () => {
    test("each call and toHtml() succeed or refuse with a documented error, within 2 seconds", async () => {
        await fc.assert(
            fc.asyncProperty(call, async input => {
                const started = performance.now();

                try {
                    await runCall(input);
                } catch (error) {
                    expectDocumentedFailure(error);
                }

                const elapsed = performance.now() - started;

                // A valid deck of a few objects builds and renders in a few milliseconds; the slowest valid cases here
                // (a table with autoPage, a chart) take well under 100 ms. 2 s leaves an order of magnitude for a slow
                // CI machine without hiding a case that loops (a structural count large enough to make a grid loop).
                expect(elapsed).toBeLessThan(2000);
            }),
            { seed: SEED, numRuns: RUNS },
        );
    }, 60_000);

    test("no call writes onto what the caller passes", async () => {
        await fc.assert(
            fc.asyncProperty(call, async input => {
                // Frozen first: a write onto the caller's object throws in these strict-mode modules.
                const frozen = deepFreeze(input);
                const before = structuredClone(input);

                try {
                    await runCall(frozen);
                } catch (error) {
                    expectDocumentedFailure(error);
                }

                expect(frozen).toEqual(before);
            }),
            { seed: SEED, numRuns: RUNS },
        );
    }, 60_000);

    test("building the same deck twice gives byte-identical HTML", async () => {
        await fc.assert(
            fc.asyncProperty(call, async input => {
                const first = await outcome(input);
                const second = await outcome(input);

                expect(second).toEqual(first);
            }),
            { seed: SEED, numRuns: RUNS },
        );
    }, 60_000);
});
