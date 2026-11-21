import type PptxGenJS from "pptxgenjs";
import type { CssColor, NormalizedImage, NormalizedSlide, NormalizedSlideObject } from "@pdflapas/model";
import type { ChartExtensionInput, ChartExtensionOptions } from "./chart/types";
import { normalizeComponent } from "./components/normalize";
import type { ComponentInput, ComponentOptions } from "./components/types";
import { PPTX_DEFAULTS } from "./defaults";
import { normalizeChart } from "./normalize/chart";
import { normalizeImage } from "./normalize/image";
import { slideMarginInches } from "./normalize/margin";
import { normalizeObjectStyle } from "./normalize/object";
import { normalizeShape, normalizeShapeTextBox } from "./normalize/shape";
import { lineOutline, normalizeShapeLine, normalizeTextLine } from "./normalize/shape/line";
import { normalizeSlideNumber } from "./normalize/slideNumber";
import { innerShadowError, normalizeColor, normalizeLine } from "./normalize/style";
import { normalizeTable } from "./normalize/table";
import { autoPageTableOptions, paginateTableRows } from "./normalize/table/pagination";
import { collapseTextMargins, floorTextWrapWidth, normalizeText, type TextInput } from "./normalize/text";
import type { PageSize } from "./pageLayouts";
import type { PptxSlide } from "./pptx";
import { chartSize, renderChartExtension, renderChartExtensionSvg } from "./render/chart";
import { renderComponent } from "./render/component";
import { objectIdScope } from "./render/slide";
import type { SlideContext } from "./slideContext";
import { RENDERING_TARGETS } from "./target";
import { UnsupportedFeatureError } from "./UnsupportedFeatureError";

/** The warnings() line for a `placeholder` the slide's master does not have, which PptxGenJS ignores without a word. */
function unknownPlaceholderWarning(method: "addText" | "addImage", name: string, hasMaster: boolean): string {
    const where = hasMaster ? "the slide's master has no placeholder" : "the slide was added without a master, so it has no placeholder";

    return `${method}({ placeholder: ${JSON.stringify(name)} }): ${where} of that name, so the object is drawn from its own options, as PptxGenJS does. Name a placeholder of the master in addSlide({ masterName }), or remove \`placeholder\`.`;
}

export class PuppeteerSlide implements PptxSlide {
    /** `context` is created by the presentation that owns the slide; every field is required, so nothing falls back to a default here. */
    constructor(private readonly context: SlideContext) {
        this.pageSize = context.pageSize;
        this.background = {
            color: PPTX_DEFAULTS.slide.backgroundColor,
            type: "solid",
        };
        this.color = PPTX_DEFAULTS.slide.color;
        this.hidden = false;
        this.newAutoPagedSlides = [];
        this.bkgd = PPTX_DEFAULTS.slide.backgroundColor;
        // The defaults above are not the slide's own background: a master's shows until the caller sets one.
        this.ownBackground = false;
        // PptxGenJS's Slide constructor copies the layout's slide number props, without the setter that writes them to the master.
        this._slideNumber = context.master?.slideNumber ?? null;
    }
    pageSize: PageSize;
    private _background: PptxGenJS.BackgroundProps = { color: PPTX_DEFAULTS.slide.backgroundColor };
    private _bkgd: string = PPTX_DEFAULTS.slide.backgroundColor;
    private _color: string = PPTX_DEFAULTS.slide.color;
    private _hidden: boolean = PPTX_DEFAULTS.slide.hidden;
    private _slideNumber: PptxGenJS.SlideNumberProps | null = null;
    /** Whether the caller set a background color or picture, which replaces the master's as the slide's <p:bg> does. */
    private ownBackground = false;
    /** The background picture, checked when it is set; drawn under every object over the whole slide. */
    private backgroundImage: NormalizedImage | undefined;
    newAutoPagedSlides: PptxGenJS.PresSlide[];
    private objectCounts: Record<string, number> = {};
    /** Title placeholder text in the theme's head font, by its index in `model.objects`, and how to draw it in another. */
    private headFontTexts: Array<{ index: number; fontFace: string; draw: (fontFace: string) => NormalizedSlideObject }> = [];
    /** Normalized, JSON-serializable record of this slide; toHtml() renders the slide from it alone. */
    readonly model: NormalizedSlide = {
        background: normalizeColor(PPTX_DEFAULTS.slide.backgroundColor),
        color: normalizeColor(PPTX_DEFAULTS.slide.color),
        hidden: PPTX_DEFAULTS.slide.hidden,
        objects: [],
    };

    get background(): PptxGenJS.BackgroundProps {
        return this._background;
    }

    set background(value: PptxGenJS.BackgroundProps) {
        // PptxGenJS takes the picture of a deprecated object `bkgd` set before (`src` as `path`) and ignores
        // `background.src`. A picture stays once set: a later color-only background does not remove it.
        const bkgd: unknown = this._bkgd;
        const old = typeof bkgd === "object" && bkgd !== null ? (bkgd as { data?: string; path?: string; src?: string }) : {};
        const data = old.data || value?.data;
        const path = old.src || old.path || value?.path;

        if (data || path) {
            this.backgroundImage = normalizeImage({ ...(data ? { data } : {}), ...(path ? { path } : {}) }, this.pageSize);
            this.ownBackground = true;
        }

        this._background = value;
        // PptxGenJS copies the deprecated `fill` to `color`.
        const color = value?.color ?? value?.fill;

        if (!color) {
            return;
        }

        this.model.background = normalizeColor({ ...value, color });
        this.ownBackground = true;
    }

    get bkgd(): string {
        return this._bkgd;
    }

    set bkgd(value: string) {
        this._bkgd = value;
        // An object (`{ path | data | src }`, deprecated) draws nothing until `background` is set, as in PptxGenJS.
        if (typeof value !== "string") {
            return;
        }

        this.model.background = normalizeColor(value);
        this.ownBackground = true;
    }

    get color(): string {
        return this._color;
    }

    set color(value: string) {
        this._color = value;
        this.model.color = normalizeColor(value);
    }

    get hidden(): boolean {
        return this._hidden;
    }

    set hidden(value: boolean) {
        // A hidden slide is left out of the PDF (toHtml() does not write it), as LibreOffice and PowerPoint export it.
        this._hidden = value;
        this.model.hidden = Boolean(value);
    }

    get slideNumber(): PptxGenJS.SlideNumberProps {
        // PptxGenJS returns null until a slide number is set; its declaration file leaves null out.
        return this._slideNumber as PptxGenJS.SlideNumberProps;
    }

    set slideNumber(value: PptxGenJS.SlideNumberProps) {
        this._slideNumber = value;
        this.context.setMasterSlideNumber(value);
    }

    /**
     * The slide's model as getModel() returns it: the background picture first, stretched over `pageSize` (the
     * presentation's), then its master's objects, then the slide's, then its number, which PptxGenJS writes last.
     * `masterSlideNumber` is the `slideNumber` set last on any slide.
     */
    toModel(pageSize: PageSize, masterSlideNumber: PptxGenJS.SlideNumberProps | null): NormalizedSlide {
        const master = this.context.master;
        const headFontFace = this.context.headFontFace();

        // PptxGenJS 4.0.1 writes a title placeholder's font as the master's +mj-lt, resolved from theme1.xml when it
        // writes the file, so text added before pres.theme set its headFontFace is drawn again in the new one.
        for (const text of this.headFontTexts) {
            if (text.fontFace !== headFontFace) {
                this.model.objects[text.index] = text.draw(headFontFace);
                text.fontFace = headFontFace;
            }
        }

        // The slide's own background replaces its master's, picture and color alike, as its <p:bg> replaces the layout's.
        const background = master?.layer && !this.ownBackground ? master.layer.backgroundLayer(pageSize) : this.backgroundLayer(pageSize);
        const masterObjects = master?.layer ? master.layer.model.objects : [];
        // The slide's number placeholder inherits what its props leave out from its layout's, then from the master's.
        const inherited = { ...masterSlideNumber, ...master?.slideNumber };
        const slideNumber = this._slideNumber ? [normalizeSlideNumber(this._slideNumber, inherited, this.context.slideIndex + 1, this.pageSize, this.context.bodyFontFace(), this.context.target, this.context.fontMetrics)] : [];

        return { ...this.model, background: background.color, objects: [...background.objects, ...masterObjects, ...this.model.objects, ...slideNumber] };
    }

    /** The slide's background color and picture over `pageSize`; a master's is read by the slides added on it. */
    private backgroundLayer(pageSize: PageSize): { color: CssColor; objects: NormalizedSlideObject[] } {
        if (!this.backgroundImage) {
            return { color: this.model.background, objects: [] };
        }

        const style = normalizeObjectStyle({ x: 0, y: 0, w: pageSize.width, h: pageSize.height }, PPTX_DEFAULTS.image, pageSize, PPTX_DEFAULTS.backgroundImage.objectName);

        // PptxGenJS writes the picture alone, so its transparent parts show white (LibreOffice), not the color.
        return { color: normalizeColor(PPTX_DEFAULTS.slide.backgroundColor), objects: [{ kind: "image", style, image: this.backgroundImage }] };
    }

    private nextObjectName(kind: string, explicit?: string): string {
        const index = this.objectCounts[kind] ?? 0;

        this.objectCounts[kind] = index + 1;

        return explicit ?? `${kind} ${index}`;
    }

    private addObject(object: NormalizedSlideObject): this {
        this.model.objects.push(object);

        return this;
    }

    private nextIdScope(): string {
        return objectIdScope(this.context.slideIndex, this.model.objects.length);
    }

    addComponent(input: ComponentInput, options: ComponentOptions): this {
        const normalized = normalizeComponent(input, options, this.pageSize, options?.objectName ?? "Component");
        const definition = this.context.components.get(normalized.input.type, normalized.input.version);

        // Application code: render it once here so its validation and render errors throw from addComponent
        // and leave the slide unchanged. The markup is dropped; toHtml() renders the model again.
        renderComponent(normalized, definition, this.nextIdScope(), this.context.getDocument);
        // The default name is only consumed after a successful render.
        const objectName = this.nextObjectName("Component", options.objectName);

        return this.addObject({ kind: "component", style: { ...normalized.style, objectName }, component: normalized.input });
    }

    // PptxGenJS declares `data: any[]`; `unknown[]` accepts every argument that does and is narrowed below.
    addChart(type: PptxGenJS.CHART_NAME | PptxGenJS.IChartMulti[], data: unknown[], options?: PptxGenJS.IChartOpts | undefined): PptxGenJS.Slide {
        // PptxGenJS treats the second argument as the shared options object for
        // IChartMulti[] calls; each mixed entry already carries its own data.
        // `as`: PptxGenJS types `data` as an array, which this call shape breaks; it is the options object here.
        const chartOptions = Array.isArray(type) && !Array.isArray(data) ? (data as unknown as PptxGenJS.IChartOpts) : (options ?? {});
        const style = normalizeObjectStyle(
            {
                x: chartOptions.x,
                y: chartOptions.y,
                w: chartOptions.w,
                h: chartOptions.h,
                objectName: chartOptions.objectName,
            },
            PPTX_DEFAULTS.chart,
            this.pageSize,
            this.nextObjectName("Chart", chartOptions.objectName),
        );
        // PptxGenJS writes `<c:holeSize>`, but not every target reads it (readsDoughnutHoleSize).
        const holeSize = RENDERING_TARGETS[this.context.target].readsDoughnutHoleSize ? chartOptions.holeSize : undefined;
        const normalized = normalizeChart(type, Array.isArray(data) ? (data as PptxGenJS.OptsChartData[]) : [], { ...chartOptions, holeSize, legendFontFace: chartOptions.legendFontFace ?? this.context.bodyFontFace() });
        // PptxGenJS always writes a marker outline, but not every target draws it (drawsChartMarkerOutline).
        const chart = RENDERING_TARGETS[this.context.target].drawsChartMarkerOutline ? normalized : { ...normalized, lineSymbolLineSize: 0 };

        chartSize(style);

        return this.addObject({ kind: "chart", style, chart });
    }

    addChartEx(input: ChartExtensionInput, options: ChartExtensionOptions): PuppeteerSlide {
        const style = normalizeObjectStyle(options, PPTX_DEFAULTS.chart, this.pageSize, this.nextObjectName("Chart", options.objectName));
        const svg = renderChartExtensionSvg(input, style);
        const altText = options.altText ?? "";

        // Parses the caller's SVG so a non-SVG root or an external reference throws here; the markup is dropped.
        renderChartExtension(svg, style, altText, this.nextIdScope());

        return this.addObject({ kind: "chartExtension", style, svg, altText });
    }

    addImage(input: PptxGenJS.ImageProps): PptxGenJS.Slide {
        const placeholder = input.placeholder ? this.context.master?.placeholders.get(input.placeholder) : undefined;

        if (input.placeholder && !placeholder) {
            this.context.warn(unknownPlaceholderWarning("addImage", input.placeholder, this.context.master !== null));
        }

        // slideObjectToXml places a picture that fills a placeholder at the placeholder's x and y and keeps its own size;
        // the placeholder's w and h are the `sizing` box when `sizing` names none.
        const area = placeholder?.options;
        const options: PptxGenJS.ImageProps = area
            ? {
                  ...input,
                  x: area.x ?? input.x,
                  y: area.y ?? input.y,
                  ...(input.sizing ? { sizing: { ...input.sizing, w: input.sizing.w || (area.w ?? input.sizing.w), h: input.sizing.h || (area.h ?? input.sizing.h) } } : {}),
              }
            : input;
        const image = normalizeImage(options, this.pageSize);
        const { transparency: _imageTransparency, ...opaqueOptions } = options;
        const geometryOptions = options.sizing ? { ...opaqueOptions, w: options.sizing.w, h: options.sizing.h } : opaqueOptions;
        // PptxGenJS writes a picture's inner shadow well-formed, but not every target draws it (drawsPictureInnerShadow).
        const shadow = options.shadow?.type === "inner" && !RENDERING_TARGETS[this.context.target].drawsPictureInnerShadow ? undefined : options.shadow;
        const style = normalizeObjectStyle({ ...geometryOptions, shadow }, PPTX_DEFAULTS.image, this.pageSize, this.nextObjectName("Image", options.objectName), {
            shadow: true,
        });

        return this.addObject({ kind: "image", style, image });
    }

    addMedia(_options: PptxGenJS.MediaProps): PptxGenJS.Slide {
        throw new UnsupportedFeatureError("addMedia", "A PDF cannot play audio or video; add a poster frame with addImage() and link the media through its `hyperlink` option.");
    }

    addNotes(_notes: string): PptxGenJS.Slide {
        // PptxGenJS writes notes to the notes page only; the slide, and so a PDF of it, never shows them.
        return this;
    }

    addShape(shapeName: PptxGenJS.SHAPE_NAME | "custGeom", options?: PptxGenJS.ShapeProps | undefined): PptxGenJS.Slide {
        const shapeOptions = options ?? {};
        const style = normalizeObjectStyle(shapeOptions, PPTX_DEFAULTS.shape, this.pageSize, this.nextObjectName("Shape", shapeOptions.objectName ?? shapeOptions.shapeName), { fill: true, shadow: true });

        style.line = normalizeShapeLine(shapeOptions, line => normalizeLine(line, lineOutline(shapeName, style.geometry.width ?? 0, style.geometry.height ?? 0)));
        const shape = normalizeShape(shapeName, shapeOptions, style.geometry.width ?? 0, style.geometry.height ?? 0, this.pageSize);

        return this.addObject({ kind: "shape", style, shape });
    }

    addTable(tableRows: PptxGenJS.TableRow[], options?: PptxGenJS.TableProps | undefined): PptxGenJS.Slide {
        const tableOptions = options ?? {};

        this.newAutoPagedSlides = [];
        if (tableOptions.autoPage) {
            // addTableDefinition and getSlidesForTableRows take a master's margin over the table's `slideMargin`.
            const master = this.context.master;
            const layoutOptions: PptxGenJS.TableProps & { slideMargin?: PptxGenJS.Margin } = master ? { ...tableOptions, slideMargin: master.margin } : tableOptions;
            const pageOptions = autoPageTableOptions(tableRows, layoutOptions, this.pageSize);
            const pages = paginateTableRows(tableRows, pageOptions, this.pageSize, master ? slideMarginInches(master.margin)[0] : PPTX_DEFAULTS.slide.marginIn);
            const continuationSlides: PuppeteerSlide[] = [];

            for (const [index, page] of pages.entries()) {
                const target = index === 0 ? this : this.context.continuationSlide(index);

                if (index > 0) {
                    continuationSlides.push(target);
                }

                target.addTable(page.rows, {
                    ...pageOptions,
                    // PptxGenJS 4.0.1 addTableDefinition (step 6 C) adds each page with the caller's options, so its
                    // rows keep the caller's rowH or h and otherwise size to their content.
                    y: page.y,
                    autoPage: false,
                });
            }

            this.newAutoPagedSlides = continuationSlides;

            return this;
        }

        const table = normalizeTable(tableRows, { ...tableOptions, fontFace: tableOptions.fontFace ?? this.context.bodyFontFace() }, this.pageSize, this.context.target, this.context.fontMetrics, {
            table: tableOptions,
            shared: this.context.sharedRunOptions,
        });
        const style = normalizeObjectStyle(tableOptions, { x: PPTX_DEFAULTS.table.x, y: PPTX_DEFAULTS.table.y }, this.pageSize, this.nextObjectName("Table", tableOptions.objectName));
        // A table is as wide and tall as its grid, whatever its frame (`w`, `h`) says; set in pixels, as an inch value of
        // 100 or more would read as EMU.
        // shortcut: addTableDefinition (step 4) reads an `x` or `y` below 20 EMU (under 0.00003in) as inches again, which
        // this does not; reproduce it with the table's `w` (addTableSize) if a deck relies on it.
        const tableStyle = { ...style, geometry: { ...style.geometry, width: table.width, height: table.height } };

        return this.addObject({ kind: "table", style: tableStyle, table });
    }

    addText(text: TextInput, options?: PptxGenJS.TextPropsOptions | undefined): PptxGenJS.Slide {
        let textOptions = options ?? {};
        let fillsPlaceholder = false;
        let placeholderSpaceBefore = {};
        let headFont = false;

        // PptxGenJS 4.0.1 addTextDefinition (A.3): text that names one of its master's placeholders takes the
        // placeholder's options over its own, and slideObjectToXml writes it with no anchor and, for a string, no
        // color. What neither sets comes from the master's titleStyle or bodyStyle (PPTX_DEFAULTS.placeholder).
        if (textOptions.placeholder) {
            const placeholder = this.context.master?.placeholders.get(textOptions.placeholder);

            if (!placeholder) {
                this.context.warn(unknownPlaceholderWarning("addText", textOptions.placeholder, this.context.master !== null));
            }

            fillsPlaceholder = placeholder !== undefined;
            const merged = { ...textOptions, ...placeholder?.options };

            // The titleStyle's +mj-lt; not a box option, as PptxGenJS writes no typeface onto the runs (RunWriting).
            headFont = placeholder?.title === true && !merged.fontFace;
            const typeStyle = placeholder?.title ? PPTX_DEFAULTS.placeholder.title : PPTX_DEFAULTS.placeholder.body;
            const fontSize = merged.fontSize ?? typeStyle.fontSizePt;

            // Each paragraph takes the style's space before from the size of its own runs (normalizeText).
            placeholderSpaceBefore = placeholder ? { placeholderSpaceBeforeRatio: typeStyle.spaceBeforeFontRatio } : {};

            textOptions = {
                ...merged,
                valign: merged.valign ?? PPTX_DEFAULTS.placeholder.valign,
                ...(typeof text === "string" && !merged.color ? { color: PPTX_DEFAULTS.placeholder.color } : {}),
                ...(placeholder
                    ? {
                          fontSize,
                          align: merged.align ?? typeStyle.align,
                      }
                    : {}),
            };
        }

        if (textOptions.shadow?.type === "inner") {
            throw innerShadowError(textOptions.shape ? `shape "${textOptions.shape}"` : "a text box");
        }

        // Text transparency is a run color property, not object opacity.
        const { transparency: _textTransparency, ...boxOptions } = textOptions;
        const textStyle = normalizeObjectStyle(boxOptions, PPTX_DEFAULTS.text, this.pageSize, this.nextObjectName("Text", textOptions.objectName), {
            fill: true,
            shadow: true,
        });

        textStyle.line = normalizeTextLine(textOptions, lineOutline(textOptions.shape ?? "", textStyle.geometry.width ?? 0, textStyle.geometry.height ?? 0));
        // PptxGenJS 4.0.1 addTextDefinition (cleanOpts) gives the box, and each of its runs without one, the box's color,
        // else the slide's, else black, and writes the box's options onto the runs as it writes the slide (RunWriting).
        // shortcut: PptxGenJS writes slide by slide, and PdfLapas reads what an earlier box wrote onto a run's options in
        // the order of the addText() and addTable() calls; they differ only when a deck adds to a slide after a later
        // one with options both share. Normalize at write time if a deck needs that order.
        const box = { ...textOptions, color: textOptions.color || this.color || PPTX_DEFAULTS.text.color };
        const color = this.color;
        // Drawn in `fontFace` where the text sets none; a title placeholder's is drawn again when the head font changes
        // (toModel).
        // shortcut: drawing again reads the run options later boxes wrote meanwhile (RunWriting), as the order note above
        // says; normalize at write time if a deck needs both.
        const draw = (fontFace: string): NormalizedSlideObject => {
            const normalizedText = normalizeText(text, { color, ...textOptions, fontFace: textOptions.fontFace ?? fontFace }, this.context.target, this.context.fontMetrics, this.context.bodyFontFace(), this.context.bodyFontFace(), {
                box,
                shared: this.context.sharedRunOptions,
                ...placeholderSpaceBefore,
            });
            // A filled placeholder has no space before its first paragraph: LibreOffice's editeng spaces a first paragraph
            // only in plain text boxes (EEControlBits::ULSPACEFIRSTPARA), and PowerPoint in none (spacesFirstParagraph).
            const [firstParagraph, ...laterParagraphs] = normalizedText.paragraphs;
            const placedText = fillsPlaceholder && firstParagraph ? { ...normalizedText, paragraphs: [{ ...firstParagraph, spaceBefore: 0 }, ...laterParagraphs] } : normalizedText;
            const height = textStyle.geometry.height ?? 0;
            const width = textStyle.geometry.width ?? 0;

            if (!textOptions.shape) {
                return { kind: "text", style: textStyle, text: collapseTextMargins(floorTextWrapWidth(placedText, width, this.context.target), height) };
            }

            // `as`: PptxGenJS writes a text box's shape from its own options object, so normalizeShape reads them as ShapeProps.
            const shapeOptions = textOptions as unknown as PptxGenJS.ShapeProps;
            const shape = normalizeShape(textOptions.shape, shapeOptions, width, height, this.pageSize);

            return { kind: "text", style: textStyle, text: collapseTextMargins(floorTextWrapWidth(normalizeShapeTextBox(placedText, shape), width, this.context.target), height), shape };
        };
        const fontFace = headFont ? this.context.headFontFace() : this.context.bodyFontFace();

        if (headFont) {
            this.headFontTexts.push({ index: this.model.objects.length, fontFace, draw });
        }

        return this.addObject(draw(fontFace));
    }
}
