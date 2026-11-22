import type { Inches, Px } from "./units";

declare const color: unique symbol;

/**
 * A color that is safe to write into CSS or SVG as is. Only the PptxGenJS color normalizer (`colorToCSS`) makes one,
 * in exactly these forms: `#RRGGBB` (the caller's six hex digits, case kept), `rgba(r, g, b, a)` (channels 0 to 255,
 * alpha in 0 to 1, below 1) for a color with transparency, or `transparent`. Type-only: the JSON holds a plain string.
 */
export type CssColor = string & { readonly [color]: "css-color" };

declare const dataUri: unique symbol;

/**
 * A resolved image source, `data:image/…;base64,…`: the only image source the HTML profile writes as `src`. Only the
 * PptxGenJS image normalizer (`imageDataUri`, `dataUri`) and its check on what `resolveImage` returns make one.
 * Type-only: the JSON holds a plain string.
 */
export type DataUri = string & { readonly [dataUri]: "data-uri" };

/** A page size as callers write it, in plain inches: what the public API takes. */
export type PageSizeInput = {
    width: number;
    height: number;
    name: string;
};

/** A page size in the model: checked, in branded inches. Assignable to PageSizeInput, so it can be passed back in. */
export type PageSize = {
    width: Inches;
    height: Inches;
    name: string;
};

export type ComponentValue = null | boolean | number | string | ComponentValue[] | { [key: string]: ComponentValue };
export type ComponentProps = { [key: string]: ComponentValue };

/** Names and versions are application-owned contracts, independent of document schema versions. */
export type ComponentInput = {
    type: string;
    version: number;
    props: ComponentProps;
};

export type Coord = number | `${number}%`;

/** The object's frame on the slide, from the slide's top left corner. */
export type NormalizedGeometry = {
    x: Px;
    y: Px;
    width?: Px;
    height?: Px;
};

export type NormalizedFill = {
    visible: boolean;
    color: CssColor;
};

export type NormalizedLine = {
    visible: boolean;
    color: CssColor;
    width: number;
    style: "solid" | "dashed" | "dotted";
    dashArray?: string;
    beginArrow?: "none" | "arrow" | "diamond" | "oval" | "stealth" | "triangle";
    endArrow?: "none" | "arrow" | "diamond" | "oval" | "stealth" | "triangle";
    /** Resolved head outlines: closed SVG path data (px) with the line end at the origin and +x pointing out of the line, filled in the line color. */
    beginHeadPath?: string;
    endHeadPath?: string;
};

export type NormalizedShadow = {
    visible: boolean;
    color: CssColor;
    blur: number;
    offsetX: number;
    offsetY: number;
    inset: boolean;
};

export type NormalizedTransform = {
    rotation: number;
    flipH: boolean;
    flipV: boolean;
    opacity: number;
};

export type NormalizedObjectStyle = {
    geometry: NormalizedGeometry;
    transform: NormalizedTransform;
    fill?: NormalizedFill;
    line?: NormalizedLine;
    shadow?: NormalizedShadow;
    objectName: string;
};

export type NormalizedTextDirection = "horizontal" | "vertical" | "vertical270" | "stacked";

/** An underline band in em of the run's font size. */
export type NormalizedUnderlineBand = {
    /** The band's top above the baseline, negative below (the sign of `post.underlinePosition`). */
    position: number;
    /** Its height, above 0. */
    thickness: number;
};

export type NormalizedTextDecoration = {
    underline: boolean;
    underlineStyle: "solid" | "double" | "dashed" | "dotted" | "wavy";
    underlineColor?: CssColor;
    /**
     * The underline is the face's `post` table band (underlineThickness tall, its top underlinePosition below the
     * baseline), drawn over descenders; omitted: the renderer's own underline.
     */
    underlineFromFont?: boolean;
    /** With `underlineFromFont`: the band to draw in place of the face's `post` band; omitted: the `post` band. */
    underlineBand?: NormalizedUnderlineBand;
    /**
     * PptxGenJS's `heavy` underline, with `underlineFromFont`: a band 1.5 times the `post` band's thickness, from
     * one thickness above its top to half one below its bottom, as PowerPoint draws it.
     */
    underlineHeavy?: boolean;
    strike: boolean;
    doubleStrike: boolean;
};

export type NormalizedTextLink = {
    href: string;
    tooltip?: string;
    slide?: number;
};

export type NormalizedTextRun = {
    text: string;
    fontFace: string;
    /** Faces a character `fontFace` lacks is drawn in, the first that has it, before the renderer's own fallback; omitted: only the renderer's. */
    fallbackFaces?: readonly string[];
    fontSize: number;
    color: CssColor;
    bold: boolean;
    italic: boolean;
    decoration: NormalizedTextDecoration;
    /** What the run asked for; `fontSize` and `baselineOffset` already hold the resulting size and shift. */
    verticalAlign: "baseline" | "sub" | "super";
    /** Baseline shift in px, positive up. */
    baselineOffset: number;
    /** Extra advance after each character in px: `charSpacing`, plus on the PowerPoint target its advance grid's rounding. */
    characterSpacing: number;
    /**
     * The extra advance in px inside each word of `text`, in order, in place of `characterSpacing`, which then spaces
     * only its spaces and tabs: the PowerPoint target's advance grid word by word. A word is a stretch of `text`
     * without a space (U+0020) or tab. Omitted: `characterSpacing` spaces every character.
     */
    wordCharacterSpacing?: readonly number[];
    /**
     * The extra advance in px on each stretch of spaces between words of `text` (a tab ends one), in order, in place
     * of `characterSpacing`: the PowerPoint target's pairs kerned on a space (Arial's " T") and spaces ending off its grid,
     * with `wordCharacterSpacing`.
     * Omitted: `characterSpacing` spaces them.
     */
    spaceCharacterSpacing?: readonly number[];
    /** `false`: drawn without the font's pair kerning (PowerPoint: below the run's `kern` size, or kerned through its letter-spacing); omitted: kerned. */
    kerning?: false;
    /** On a bold run: the stroke width, as a fraction of `fontSize`, that emboldens a face with no bold; omitted: the renderer's own. */
    syntheticBold?: number;
    /** On an italic run: the shear that slants a face with no italic (x per unit of height); omitted: the renderer's own. */
    syntheticItalic?: number;
    /** On a run whose text has U+2011: "hyphen-minus" draws it with the face's U+002D glyph, still unbreakable; omitted: as any character. */
    nonBreakingHyphen?: "hyphen-minus";
    highlight?: CssColor;
    outline?: { color: CssColor; width: number };
    glow?: NormalizedShadow;
    shadow?: NormalizedShadow;
    language: string;
    link?: NormalizedTextLink;
    softBreakBefore: boolean;
};

export type NormalizedTextBullet = {
    kind: "bullet" | "number";
    marker: string;
    indent: number;
    level: number;
    start: number;
    numberType?: string;
};

export type NormalizedTextParagraph = {
    runs: NormalizedTextRun[];
    align: "left" | "center" | "right" | "justify";
    /** With `align: "justify"`: `"inter-word"` widens only a line's spaces (CSS `text-justify`); absent, the renderer's `auto`. */
    textJustify?: "inter-word";
    rtl: boolean;
    /**
     * Each line's height and its baseline's distance below the line top, in px (a font-independent line box, as the
     * target draws it). `markBaseline` replaces `baseline` on the paragraph's last line and `breakBaseline` on a line
     * that ends in a soft break (`softBreakBefore` on the next run); absent, those lines use `baseline` too.
     */
    lineBox: { height: number; baseline: number; markBaseline?: number; breakBaseline?: number };
    /** Adjacent spaces collapse to the larger (CSS margins); a producer that wants them added moves the space after here. */
    spaceBefore: number;
    /** 0 on a text box's last paragraph: PowerPoint and LibreOffice leave its space after out of the text height. */
    spaceAfter: number;
    /** Explicit stops, `position` in px from the text area's left edge, in increasing order. */
    tabStops: ReadonlyArray<{ position: number; alignment: "left" | "right" | "center" | "decimal" }>;
    /** Spacing of the default stops after the last explicit one, in px. */
    defaultTabStop: number;
    bullet?: NormalizedTextBullet;
    /** Without a bullet: where every line starts, in px from the text area's left edge; absent, 0. */
    marginLeft?: number;
};

export type NormalizedTextBox = {
    paragraphs: NormalizedTextParagraph[];
    margin: [number, number, number, number];
    verticalAlign: "top" | "middle" | "bottom";
    wrap: boolean;
    fit: "none" | "shrink" | "resize";
    direction: NormalizedTextDirection;
    rtl: boolean;
};

export type NormalizedTableBorder = {
    visible: boolean;
    color: CssColor;
    width: number;
    style: "solid" | "dashed";
    /** SVG stroke-dasharray in px for a dashed border, resolved from the DrawingML preset dash. */
    dashArray?: string;
};

export type NormalizedTableCell = {
    text: NormalizedTextBox;
    colspan: number;
    rowspan: number;
    fill?: NormalizedFill;
    borders: [NormalizedTableBorder, NormalizedTableBorder, NormalizedTableBorder, NormalizedTableBorder];
};

export type NormalizedTableRow = {
    height?: Px;
    cells: NormalizedTableCell[];
};

export type NormalizedTable = {
    width: Px;
    height?: Px;
    columns: Px[];
    rows: NormalizedTableRow[];
    /**
     * Px a row without a height (when the table has none either) is taller than its tallest cell's text and
     * margins, as free space of each cell's text area, where the cell's vertical anchor places the text.
     */
    contentRowSlack?: Px;
};

/** `contain`/`cover`: the image's box inside the object's frame; `crop`: the whole image's size and the offset of the part the frame shows. */
export type NormalizedImageSizing = { type: "stretch" } | { type: "contain" | "cover"; x: Px; y: Px; width: Px; height: Px } | { type: "crop"; offsetX: Px; offsetY: Px; width: Px; height: Px };

export type NormalizedImageLink = {
    href: string;
    tooltip?: string;
    slide?: number;
};

/** A `data` image is inlined at normalization; a `path` or `remote` one keeps the caller's string until `toHtml()` resolves it. */
export type NormalizedImageSource = { source: DataUri; sourceKind: "data" } | { source: string; sourceKind: "path" | "remote" };

export type NormalizedImage = NormalizedImageSource & {
    sizing: NormalizedImageSizing;
    rounding: boolean;
    opacity: number;
    altText: string;
    link?: NormalizedImageLink;
};

export type NormalizedShapeLink = {
    href: string;
    tooltip?: string;
    slide?: number;
};

export type NormalizedShapeGeometry =
    | { kind: "rect"; radius: Px }
    | { kind: "ellipse" }
    | { kind: "line"; inverse: boolean }
    | {
          kind: "path";
          data: string;
          transform?: string;
          faces?: ReadonlyArray<{ data: string; fillModifier?: "darken" | "darkenLess" | "lighten" | "lightenLess" }>;
          outlineData?: string;
      };

export type NormalizedShape = {
    name: string;
    width: Px;
    height: Px;
    geometry: NormalizedShapeGeometry;
    textInsets?: [number, number, number, number];
    link?: NormalizedShapeLink;
};

export type NormalizedChartType = "area" | "bar" | "bar3D" | "bubble" | "doughnut" | "line" | "mixed" | "pie" | "radar" | "scatter";

export type NormalizedChartSeries = {
    type: NormalizedChartType;
    name: string;
    labels: string[];
    values: number[];
    sizes?: number[];
    colorIndex: number;
    varyColors: boolean;
    valueAxisIndex: 0 | 1;
    /**
     * How the series stacks with the other series of its type on its value axis (its chart element's `<c:grouping>`);
     * absent: it is drawn on its own.
     */
    stack?: "stacked" | "percentStacked";
    /** Each point's fill in place of the series color (a `<c:dPt>` without a fill); absent: the series color. */
    automaticPointFill?: CssColor;
    /** Each point's outline color, by point index (a `<c:dPt>` `<a:ln>`); absent: no outline. */
    pointLineColors?: CssColor[];
    /** The series' own data labels (a combo chart's chart type options); absent: the chart's. */
    dataLabels?: NormalizedChartDataLabels;
};

export type NormalizedChartValueAxis = {
    minimum: number;
    maximum: number;
    hidden: boolean;
    showTitle: boolean;
    title: string;
    titleColor: CssColor;
    titleFontFace: string;
    titleFontSize: number;
    titleRotate?: number;
    labelColor: CssColor;
    labelFontFace: string;
    labelFontSize: number;
    labelBold: boolean;
    labelItalic: boolean;
    labelRotate: number;
    labelPosition: "none" | "low" | "high" | "nextTo";
    labelFormatCode: string;
    lineColor: CssColor;
    lineWidth: number;
    lineStyle: "solid" | "dash" | "dot";
    lineVisible: boolean;
    majorTickMark: "none" | "inside" | "outside" | "cross";
    minorTickMark: "none" | "inside" | "outside" | "cross";
    majorUnit?: number;
    /** Whether the axis runs from its maximum to its minimum (`valAxisOrientation: "maxMin"`); absent: minimum first. */
    reversed?: boolean;
    /**
     * On a bar3D chart, whose step LibreOffice picks from the drawn length of the axis: the data's low and high ends (0
     * included) and whether the caller set the minimum and the maximum; absent: `minimum` to `maximum` as they are.
     * `labelHeightsPerStep` is how many label heights of drawn axis each step needs; absent: one.
     */
    automaticScale?: { low: number; high: number; fixedMinimum: boolean; fixedMaximum: boolean; labelHeightsPerStep?: number };
    logScaleBase?: number;
    gridLineColor: CssColor;
    gridLineWidth: number;
    gridLineStyle: "solid" | "dash" | "dot" | "none";
    gridLineCap: "flat" | "round" | "square";
};

/** The data labels of a chart or a series: what each label shows, where and in which text style. */
export type NormalizedChartDataLabels = {
    showValue: boolean;
    /** The category name, on pie and doughnut slices; a scatter point's own label (its series' `labels`). */
    showCategory: boolean;
    /** The series name; absent: not shown. */
    showSeriesName?: boolean;
    /** What joins the series name and the value (the rendering target's); absent: ", ". */
    fieldSeparator?: string;
    /** Whether a bubble's value label prints its size in place of its y value; absent: the y value. */
    valueIsBubbleSize?: boolean;
    /** The share of the total, on pie and doughnut slices. */
    showPercent: boolean;
    position: "b" | "bestFit" | "ctr" | "inBase" | "inEnd" | "l" | "outEnd" | "r" | "t";
    valueFormat: string;
    percentFormat: string;
    color: CssColor;
    fontFace: string;
    fontSize: number;
    bold: boolean;
    italic: boolean;
};

/** The category axis of a bar, line or area chart: its labels, title and gridlines. */
export type NormalizedChartCatAxis = {
    showTitle: boolean;
    title: string;
    titleColor: CssColor;
    titleFontFace: string;
    titleFontSize: number;
    titleRotate?: number;
    labelColor: CssColor;
    labelFontFace: string;
    labelFontSize: number;
    labelBold: boolean;
    labelItalic: boolean;
    gridLineColor: CssColor;
    gridLineWidth: number;
    gridLineStyle: "solid" | "dash" | "dot" | "none";
    gridLineCap: "flat" | "round" | "square";
};

export type NormalizedChart = {
    type: NormalizedChartType;
    series: NormalizedChartSeries[];
    colors: CssColor[];
    colorOpacity: number;
    showLegend: boolean;
    legendPosition: "b" | "l" | "r" | "t" | "tr";
    /**
     * Where a top-right legend goes: in its band beside the plot area, or in a band above it that leaves the plot its
     * full width, its first key centered firstKeyCenterEm legend font sizes below the chart's top edge, one entry every
     * pitchEm, and the band ending bandBelowEm below the last key's center.
     */
    topRightLegend: { kind: "beside" } | { kind: "above"; firstKeyCenterEm: number; pitchEm: number; bandBelowEm: number };
    showTitle: boolean;
    title: string;
    fontFace: string;
    fontSize: number;
    titleFontFace: string;
    titleFontSize: number;
    titleBold: boolean;
    titleColor: CssColor;
    legendFontFace: string;
    legendFontSize: number;
    legendColor: CssColor;
    barDirection: "bar" | "col";
    grouping: "clustered" | "stacked" | "percentStacked" | "standard";
    barGapWidthPercent: number;
    barGapDepthPercent: number;
    bar3DShape: "box" | "cylinder" | "cone" | "coneToMax" | "pyramid" | "pyramidToMax";
    perspective3D: number;
    rightAngleAxes3D: boolean;
    rotationX3D: number;
    rotationY3D: number;
    gridLineColor: CssColor;
    gridLineWidth: number;
    axisLineVisible: boolean;
    categoryMajorTickMark: "none" | "inside" | "outside" | "cross";
    valueAxisMinimum: number;
    valueAxisMaximum: number;
    valueAxes: NormalizedChartValueAxis[];
    categoryAxis: NormalizedChartCatAxis;
    doughnutHoleSize: number;
    firstSliceAngle: number;
    radarStyle: "standard" | "marker" | "filled";
    lineSmooth: boolean;
    lineSymbol: "circle" | "dash" | "diamond" | "dot" | "none" | "square" | "triangle";
    lineSymbolSize: number;
    /** The marker outline's color; absent: the series color. */
    lineSymbolLineColor?: CssColor;
    /** The marker outline's width in points; 0 draws none. */
    lineSymbolLineSize: number;
    lineSize: number;
    dataLabels: NormalizedChartDataLabels;
    plotArea: NormalizedChartPlotArea;
    /** The inner plot rectangle as fractions of the chart frame; absent for the automatic layout. */
    plotLayout?: NormalizedChartLayout;
    /** Whether the series are drawn from the last to the first, so the first is in front; absent: in series order. */
    firstSeriesInFront?: boolean;
    /**
     * Whether the legend lists the only series' points, each by its category label and keyed in its point's color,
     * in place of the series; absent: the series.
     */
    legendListsPoints?: boolean;
    /** Whether a bar3D chart with one unstacked series names it at the depth axis; absent: no depth label for it. */
    namesSingleDepthSlot?: boolean;
    altText: string;
};

export type NormalizedChartPlotArea = {
    fill: NormalizedFill;
    borderVisible: boolean;
    borderColor: CssColor;
    /** Points. */
    borderWidth: number;
};

export type NormalizedChartLayout = { x: number; y: number; w: number; h: number };

/**
 * One slide object as recorded by the PptxGenJS-compatible API, after option
 * normalization. Every variant is plain JSON data. `component` keeps only the
 * registry key, version and JSON props; rendering it needs the application's
 * registered definition. `chartExtension` keeps the SVG produced from the
 * caller's renderer input, because ECharts options may contain functions.
 */
export type NormalizedSlideObject =
    | { kind: "text"; style: NormalizedObjectStyle; text: NormalizedTextBox; shape?: NormalizedShape }
    | { kind: "shape"; style: NormalizedObjectStyle; shape: NormalizedShape }
    | { kind: "image"; style: NormalizedObjectStyle; image: NormalizedImage }
    | { kind: "table"; style: NormalizedObjectStyle; table: NormalizedTable }
    | { kind: "chart"; style: NormalizedObjectStyle; chart: NormalizedChart }
    | { kind: "chartExtension"; style: NormalizedObjectStyle; svg: string; altText: string }
    | { kind: "component"; style: NormalizedObjectStyle; component: ComponentInput };

export type NormalizedSlide = {
    /** CSS colors, as applied to the slide container. */
    background: CssColor;
    color: CssColor;
    /** Left out of the HTML and the PDF; the other slides keep their number. */
    hidden: boolean;
    /**
     * In paint order. PptxGenJS's background picture is the first object (an image named "Background" over the whole
     * slide) and its slide number the last (a text box named "Slide Number").
     */
    objects: NormalizedSlideObject[];
};

/**
 * Document properties, as PptxGenJS writes them to docProps/core.xml (title, subject, author as dc:creator, revision)
 * and docProps/app.xml (company), plus the PDF-only ones PptxGenJS has no field for (`PuppeteerGenOptions.pdfMetadata`).
 * An absent or empty one is not written: no element in the HTML, no entry in the PDF.
 */
export type PresentationMetadata = {
    title?: string;
    author?: string;
    subject?: string;
    company?: string;
    revision?: string;
    keywords?: string;
    /** The PDF /Creator; renderers write `PDF_PRODUCER` without one. */
    creator?: string;
    /** The PDF /Producer; renderers write `PDF_PRODUCER` without one. */
    producer?: string;
    /** UTC, as `Date.prototype.toISOString()` writes it (`2026-11-19T12:34:56.000Z`). */
    creationDate?: string;
    /** UTC, as `creationDate`. */
    modificationDate?: string;
};

export type NormalizedPresentation = {
    pageSize: PageSize;
    metadata?: PresentationMetadata;
    slides: NormalizedSlide[];
};
