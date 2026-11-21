import type PptxGenJS from "pptxgenjs";
import type { NormalizedPresentation, PresentationMetadata } from "@pdflapas/model";
import { ComponentRegistry } from "./components/registry";
import type { ComponentDefinition, ComponentProps } from "./components/types";
import { type FontMetricsLookup, type FontMetricsResult, type FontWinMetrics, type GlyphAdvancesResult, PPTX_DEFAULTS, PPTXGENJS_VERSION, type UnreadableFont } from "./defaults";
import { DEFAULT_RENDERING_TARGET, RENDERING_TARGETS, type RenderingTarget } from "./target";
import { ALIGN_H, ALIGN_V, CHART_TYPE, DEPRECATED_CHART_TYPE, DEPRECATED_SCHEME_COLOR, DEPRECATED_SHAPE_TYPE, OUTPUT_TYPE, SCHEME_COLOR, SHAPE_TYPE } from "./enums";
import { type ImageSourceOptions, imageResolver, resolveRemoteImageSource } from "./normalize/image";
import { fontMetricsLookup } from "./fontMetrics";
import { normalizePdfMetadata, type PdfMetadataOptions } from "./normalize/pdfMetadata";
import type { SharedRunOptions } from "./normalize/sharedRunOptions";
import { findHtmlTable, normalizeHtmlTable } from "./normalize/table/html";
import type { TableHtmlLimits } from "./normalize/table/htmlLimits";
import { paginateTableRows } from "./normalize/table/pagination";
import { convertPdfOutput, downloadPdf, resolveFileName, resolveOutputType } from "./output";
import type { PageSize, PageSizeInput } from "./pageLayouts";
import { checkedPageSize, DEFAULT_PAGE_SIZE, PPTX_LAYOUT_PAGE_SIZES } from "./pageLayouts";
import type { PptxAddSlideProps, PptxGenJSLike, PptxSectionProps, PptxSlideMasterProps, PptxTableToSlidesProps, PptxWriteBaseProps, PptxWriteFileProps, PptxWriteProps } from "./pptx";
import { PuppeteerSlide } from "./PuppeteerSlide";
import type { MasterPlaceholder, SlideContext, SlideMaster } from "./slideContext";
import { type HtmlPresentationDocument, lazyDocument, type PresentationRenderer, trustedModelToHtml } from "./render/document";
import { createBrowserDocument } from "./render/slide";

/**
 * PptxGenJS's enum objects. Each one is a plain object with the enum's keys and values, which tests/presentation.test.ts
 * compares with the pinned runtime; the declaration file's enum types cannot see that, hence the one cast.
 */
function pptxEnum<Enum>(table: object): Enum {
    return table as Enum;
}

/** Runtime adapters. The Node entry supplies jsdom, filesystem image paths and Puppeteer by default. */
export type PuppeteerGenOptions = {
    /** Turns the toHtml() document into PDF bytes for toPdf() and writeFile(); every renderer reads the HTML (#69). */
    renderer?: PresentationRenderer;
    /**
     * Resolves an addImage() path or URL to a `data:image/…;base64,` URI before serialization; toHtml() throws
     * UnresolvedImageError for any other result, such as a URL or a path.
     */
    resolveImage?: (source: string) => Promise<string>;
    /**
     * Limits and confinement of the default resolver: `timeoutMs`, `maxBytes`, `baseDir` (Node entry) and
     * `allowRemote`. Unset, any readable path and any http(s) URL resolve, as in PptxGenJS, within the default
     * limits. Cannot be combined with `resolveImage`.
     */
    imageSources?: ImageSourceOptions;
    /** Creates the documents tableToSlides() parses HTML strings into and components may build elements in, only when one is needed. */
    createDocument?: () => Document;
    /**
     * Whose rendering the PDF matches where LibreOffice and PowerPoint draw the same file differently: "libreoffice"
     * (the default) or "powerpoint". PuppeteerGenLibreOffice and PuppeteerGenPowerPoint set it.
     */
    target?: RenderingTarget;
    /**
     * head.unitsPerEm and OS/2 usWinAscent and usWinDescent of the faces a deck uses, keyed by face name
     * (case-insensitive), which the PowerPoint target places a single-spaced line's baseline from. They override
     * `resolveFontMetrics` and the built-in FONT_WIN_METRICS; a face in none keeps the font-independent baseline.
     * `readFontWinMetrics(bytes)` reads them from a font file. `pairKerning: false` says the face has no pair kerning
     * PowerPoint reads (no version 0 `kern` table), so the PowerPoint target draws it unkerned at any size; omitted, it
     * is kerned from 12pt unless its `resolveGlyphAdvances` advances carry no `kern` pairs. The LibreOffice target
     * ignores them, as LibreOffice's line box ignores the font and it kerns every run.
     */
    fontMetrics?: Readonly<Record<string, FontWinMetrics>>;
    /**
     * Looks a face's metrics up for the PowerPoint target when `fontMetrics` does not list it: `{ kind: "metrics",
     * metrics }` of the font the face resolves to, `{ kind: "unknown" }` to use FONT_WIN_METRICS, or `{ kind:
     * "unreadable", face, path, reason }` for a font file it found but could not read, which is laid out as unknown and
     * listed in `warnings()`. Called at most once per face per presentation, and never for the LibreOffice target; its
     * result is checked. The Node entry reads the installed font by default (fontconfig); `null` turns the lookup off, so
     * the HTML does not depend on the machine's fonts.
     */
    resolveFontMetrics?: ((fontFace: string) => FontMetricsResult | UnreadableFont) | null;
    /**
     * Looks up the advance widths of a face in a weight and style for the PowerPoint target, which lays text out with
     * each advance rounded to 0.125pt (RENDERING_TARGETS.powerpoint.advanceGridPt): `{ kind: "advances", advances:
     * readGlyphAdvances(bytes) }` of the font, `{ kind: "unknown" }` to keep the natural advances, or `{ kind:
     * "unreadable", face, path, reason }` as for `resolveFontMetrics`. Called at most once per face, weight and style per
     * presentation, never for the LibreOffice target; its result is checked. The Node entry reads the installed font by
     * default (fontconfig); `null`, or no resolver in the browser build, keeps every face's natural advances.
     */
    resolveGlyphAdvances?: ((fontFace: string, bold: boolean, italic: boolean) => GlyphAdvancesResult | UnreadableFont) | null;
    /**
     * PDF document properties PptxGenJS has no field for: `creator`, `producer`, `keywords`, `creationDate` and
     * `modificationDate`. Unset, Creator and Producer are "PdfLapas" and the PDF has no dates, so the same deck gives
     * the same bytes. Dates are written in UTC to the second; one date alone is used for both.
     */
    pdfMetadata?: PdfMetadataOptions;
};

// PptxGenJS 4.0.1 declares `PlaceholderType` but its instance has no such member (it reads undefined), so neither has this one.
export class PuppeteerGen implements Omit<PptxGenJSLike, "PlaceholderType"> {
    private pageSize: PageSize;
    /** The name last assigned to `layout`; undefined until then, as in PptxGenJS. */
    private layoutName: string | undefined;
    private presLayoutName: string;
    /** `defineLayout()` sizes by name, which `layout` can then select (PptxGenJS's `LAYOUTS`). */
    private readonly definedLayouts = new Map<string, PageSize>();
    private themeProps: PptxGenJS.ThemeProps | undefined;
    private rtl: boolean = PPTX_DEFAULTS.presentation.rtlMode;
    // PptxGenJS writes these to docProps/core.xml and app.xml; toHtml() writes them to <head> and the renderers to the PDF.
    // Only what the caller sets: PptxGenJS's defaults ("PptxGenJS", "PptxGenJS Presentation", "1") are its branding,
    // not the caller's, so an unset one reads "" and reaches neither the HTML nor the PDF (owner decision, #102).
    private metadata: PresentationMetadata = {};
    /** `PuppeteerGenOptions.pdfMetadata`, checked: PdfLapas's own properties, kept apart from PptxGenJS's `pres.*`. */
    private readonly pdfMetadata: PresentationMetadata;
    /** The `slideNumber` set last on any slide: PptxGenJS writes it to the master, which slide numbers inherit from. */
    private masterSlideNumber: PptxGenJS.SlideNumberProps | null = null;
    private readonly components = new ComponentRegistry();
    /** What PptxGenJS writes onto run options objects the deck shares between boxes, which its slides read and extend. */
    private readonly sharedRunOptions: SharedRunOptions = new WeakMap();
    private slides: PuppeteerSlide[] = [];
    /** defineSlideMaster() layouts by `title`; the first one defined under a title wins, as PptxGenJS finds it first. */
    private readonly masters = new Map<string, SlideMaster>();
    private readonly renderer: PresentationRenderer | undefined;
    private readonly resolveImage: (source: string) => Promise<string>;
    private readonly createDocument: () => Document;
    private readonly target: RenderingTarget;
    private readonly fontMetrics: FontMetricsLookup;
    /** What warnings() lists, each line once, in the order found. */
    private readonly warningLines = new Set<string>();
    /** The `context.document` of components checked in addComponent(), created the first time one reads it. */
    private readonly componentDocument: () => Document;

    constructor(pageSize: PageSizeInput = DEFAULT_PAGE_SIZE, options: PuppeteerGenOptions = {}) {
        this.pageSize = checkedPageSize("new PuppeteerGen(pageSize)", pageSize);
        this.presLayoutName = pageSize === DEFAULT_PAGE_SIZE ? PPTX_DEFAULTS.presentation.layouts.LAYOUT_16x9.presLayoutName : this.pageSize.name;
        this.renderer = options.renderer;
        this.resolveImage = imageResolver(options, resolveRemoteImageSource);
        this.createDocument = options.createDocument ?? createBrowserDocument;

        // The constructor's `target`, checked: a misspelled one would otherwise render silently for the default.
        if (options.target === undefined) {
            this.target = DEFAULT_RENDERING_TARGET;
        } else if (typeof options.target === "string" && Object.hasOwn(RENDERING_TARGETS, options.target)) {
            this.target = options.target;
        } else {
            throw new TypeError(
                `PuppeteerGen option \`target\` must be one of ${Object.keys(RENDERING_TARGETS)
                    .map(name => `"${name}"`)
                    .join(", ")}, got ${JSON.stringify(options.target)}`,
            );
        }

        this.pdfMetadata = normalizePdfMetadata(options.pdfMetadata);
        this.fontMetrics = fontMetricsLookup(options.fontMetrics, options.resolveFontMetrics, options.resolveGlyphAdvances, message => this.warningLines.add(message));
        this.componentDocument = lazyDocument(this.createDocument);
    }

    get version(): string {
        return PPTXGENJS_VERSION;
    }

    /**
     * PptxGenJS `layout`: LAYOUT_16x9, LAYOUT_16x10, LAYOUT_4x3, LAYOUT_WIDE or a `defineLayout()` name. The page size
     * applies to every slide, as PptxGenJS writes one slide size; a slide's percentage positions use the size it had when
     * it was added, as PptxGenJS resolves them against that slide's layout.
     */
    get layout(): string {
        // PptxGenJS returns undefined until a layout is assigned, although its declaration says string.
        return this.layoutName as string;
    }

    set layout(name: string) {
        const defined = this.definedLayouts.get(name);

        if (defined) {
            this.pageSize = defined;
            this.presLayoutName = name;
        } else if (Object.hasOwn(PPTX_LAYOUT_PAGE_SIZES, name)) {
            // `as`: hasOwn checked it is a built-in layout's key.
            const builtIn = name as keyof typeof PPTX_LAYOUT_PAGE_SIZES;

            this.pageSize = PPTX_LAYOUT_PAGE_SIZES[builtIn];
            this.presLayoutName = PPTX_DEFAULTS.presentation.layouts[builtIn].presLayoutName;
        } else {
            // PptxGenJS throws Error('UNKNOWN-LAYOUT') here; keep that code and say how to fix it.
            throw new Error(`UNKNOWN-LAYOUT: '${name}' is not a layout. Use ${Object.keys(PPTX_LAYOUT_PAGE_SIZES).join(", ")}, or call pres.defineLayout({ name: '${name}', width, height }) first.`);
        }

        this.layoutName = name;
    }

    /** PptxGenJS `presLayout`: the current page size in EMU and its PptxGenJS name (`screen16x9`, a `defineLayout()` name). */
    get presLayout(): PptxGenJS.PresLayout {
        const emuPerInch = PPTX_DEFAULTS.units.emuPerInch;

        return { name: this.presLayoutName, width: Math.round(this.pageSize.width * emuPerInch), height: Math.round(this.pageSize.height * emuPerInch) };
    }

    /**
     * PptxGenJS `theme`. `bodyFontFace` is the font of text, table cells and chart legends that name none, as PowerPoint
     * resolves their `+mn-lt` from theme1.xml. `headFontFace` only styles title placeholders, which need a slide master.
     */
    get theme(): PptxGenJS.ThemeProps {
        // PptxGenJS returns undefined until a theme is assigned, although its declaration says ThemeProps.
        return this.themeProps as PptxGenJS.ThemeProps;
    }

    set theme(value: PptxGenJS.ThemeProps) {
        for (const key of ["headFontFace", "bodyFontFace"] as const) {
            const font: unknown = value?.[key];

            if (font !== undefined && typeof font !== "string") {
                throw new TypeError(`pres.theme.${key} must be a font name such as "Arial", got ${JSON.stringify(font)}`);
            }
        }

        // PptxGenJS treats an empty font name as unset (`theme.bodyFontFace ? ... : Calibri`).
        const bodyFontFace = value?.bodyFontFace || PPTX_DEFAULTS.theme.bodyFontFace;

        const layers = [...this.slides, ...Array.from(this.masters.values(), master => master.layer)];

        if (bodyFontFace !== this.bodyFontFace() && layers.some(slide => slide.model.objects.length > 0)) {
            throw new Error(
                `pres.theme.bodyFontFace "${bodyFontFace}" was set after slide objects were added: PdfLapas resolves an object's default font when it is added, PptxGenJS when it writes the file. Set pres.theme before adding text, tables or charts.`,
            );
        }

        this.themeProps = value;
    }

    /**
     * PptxGenJS `rtlMode` writes `<p:presentation rtl="1">`, which orients PowerPoint's editing views; slides look the
     * same, so the PDF does too. Right-to-left text is the `rtlMode` option of addText().
     */
    get rtlMode(): boolean {
        return this.rtl;
    }

    set rtlMode(value: boolean) {
        this.rtl = value;
    }

    get author(): string {
        return this.metadata.author ?? "";
    }

    set author(value: string) {
        this.setMetadata("author", value);
    }

    get company(): string {
        return this.metadata.company ?? "";
    }

    set company(value: string) {
        this.setMetadata("company", value);
    }

    get revision(): string {
        return this.metadata.revision ?? "";
    }

    set revision(value: string) {
        this.setMetadata("revision", value);
    }

    get subject(): string {
        return this.metadata.subject ?? "";
    }

    set subject(value: string) {
        this.setMetadata("subject", value);
    }

    get title(): string {
        return this.metadata.title ?? "";
    }

    set title(value: string) {
        this.setMetadata("title", value);
    }

    private setMetadata(key: keyof PresentationMetadata, value: unknown): void {
        if (typeof value !== "string") {
            throw new TypeError(`pres.${key} must be a string, got ${value === null ? "null" : typeof value}; an empty string leaves it out of the PDF`);
        }

        this.metadata = { ...this.metadata, [key]: value };
    }

    get AlignH(): typeof PptxGenJS.AlignH {
        return pptxEnum<typeof PptxGenJS.AlignH>(ALIGN_H);
    }

    get AlignV(): typeof PptxGenJS.AlignV {
        return pptxEnum<typeof PptxGenJS.AlignV>(ALIGN_V);
    }

    get ChartType(): typeof PptxGenJS.ChartType {
        return pptxEnum<typeof PptxGenJS.ChartType>(CHART_TYPE);
    }

    get OutputType(): typeof PptxGenJS.OutputType {
        return pptxEnum<typeof PptxGenJS.OutputType>(OUTPUT_TYPE);
    }

    get SchemeColor(): typeof PptxGenJS.SchemeColor {
        return pptxEnum<typeof PptxGenJS.SchemeColor>(SCHEME_COLOR);
    }

    get ShapeType(): typeof PptxGenJS.ShapeType {
        return pptxEnum<typeof PptxGenJS.ShapeType>(SHAPE_TYPE);
    }

    /** @deprecated PptxGenJS's runtime alias of `ChartType`, absent from its types. */
    get charts(): typeof DEPRECATED_CHART_TYPE {
        return DEPRECATED_CHART_TYPE;
    }

    /** @deprecated PptxGenJS's runtime alias of `SchemeColor`, absent from its types. */
    get colors(): typeof DEPRECATED_SCHEME_COLOR {
        return DEPRECATED_SCHEME_COLOR;
    }

    /** @deprecated PptxGenJS's runtime alias of `ShapeType`, absent from its types. */
    get shapes(): typeof DEPRECATED_SHAPE_TYPE {
        return DEPRECATED_SHAPE_TYPE;
    }

    private bodyFontFace(): string {
        return this.themeProps?.bodyFontFace || PPTX_DEFAULTS.theme.bodyFontFace;
    }

    registerComponent<Props extends object = ComponentProps>(type: string, version: number, definition: ComponentDefinition<Props>): this {
        this.components.register(type, version, definition);

        return this;
    }

    /**
     * Set a custom page size
     */
    setPageSize(pageSize: PageSizeInput): void {
        this.pageSize = checkedPageSize("setPageSize()", pageSize);
        this.presLayoutName = this.pageSize.name;
    }

    /**
     * The normalized slide-object model that toHtml() and toPdf() render, as a detached JSON-serializable copy.
     * `component` objects keep only type, version and props; rendering them needs the registered definitions.
     */
    getModel(): NormalizedPresentation {
        return structuredClone({ pageSize: this.pageSize, metadata: { ...this.metadata, ...this.pdfMetadata }, slides: this.slides.map(slide => slide.toModel(this.pageSize, this.masterSlideNumber)) });
    }

    /**
     * What this presentation laid out differently than asked without failing, one line each: today a font file a font
     * resolver found but could not read, whose face falls back to the built-in metrics and natural advances (and can wrap
     * differently from PowerPoint). Filled as objects are laid out, so read it after toHtml(), toPdf() or getModel().
     */
    warnings(): string[] {
        return [...this.warningLines];
    }

    /** PptxGenJS `stream()`: the PDF as a Node Buffer. `compression` is ignored (zip-only). */
    async stream(_props?: PptxWriteBaseProps | undefined): Promise<string | ArrayBuffer | Blob | Uint8Array> {
        return convertPdfOutput(await this.toPdf(), "STREAM");
    }

    /** PptxGenJS `write()`: the PDF as the requested outputType, default `blob`. `compression` is ignored (zip-only). */
    async write(props?: PptxWriteProps | undefined): Promise<string | ArrayBuffer | Blob | Uint8Array> {
        const outputType = resolveOutputType(props);

        return convertPdfOutput(await this.toPdf(), outputType);
    }

    /** The complete self-contained HTML document (images inlined as data URIs). Does not launch a browser. */
    async toHtml(): Promise<HtmlPresentationDocument> {
        return trustedModelToHtml(this.getModel(), { components: this.components, resolveImage: this.resolveImage, createDocument: this.createDocument });
    }

    /** PDF bytes from the configured renderer; nothing is written to disk. */
    async toPdf(): Promise<Uint8Array> {
        if (!this.renderer) {
            throw new Error('toPdf() requires a renderer: pass { renderer } to the PuppeteerGen constructor, or use the Node entry ("@pdflapas/pptxgenjs") for the Puppeteer default');
        }

        // A renderer with only renderModel() (the model input, removed in #69) would otherwise fail as "render is not a function".
        if (typeof this.renderer.render !== "function") {
            throw new Error("toPdf() renderer must implement render(document: HtmlPresentationDocument): every renderer reads the toHtml() document, renderModel(model) was removed (#69)");
        }

        return this.renderer.render(await this.toHtml());
    }

    /** PptxGenJS `writeFile()`: in a browser, downloads the PDF as PptxGenJS downloads its .pptx; returns the file name. */
    async writeFile(props?: PptxWriteFileProps | string): Promise<string> {
        const fileName = resolveFileName(props);

        if (typeof document === "undefined" || typeof URL.createObjectURL !== "function") {
            throw new Error('writeFile() downloads the PDF in a browser, or writes it to disk from the Node entry ("@pdflapas/pptxgenjs"); this runtime has neither a document nor a filesystem, so use toPdf() or toHtml() instead');
        }

        downloadPdf(document, fileName, await this.toPdf());

        return fileName;
    }

    addSection(props: PptxSectionProps): void {
        // Sections only group slides in PowerPoint's slide list and draw nothing; keep PptxGenJS's input warnings.
        if (!props) {
            console.warn("addSection requires an argument");
        } else if (!props.title) {
            console.warn("addSection requires a title");
        }
    }

    addSlide(props?: PptxAddSlideProps | undefined): PuppeteerSlide;
    addSlide(masterName?: string | undefined): PuppeteerSlide;
    addSlide(props?: unknown): PuppeteerSlide {
        // A string is PptxGenJS's deprecated master name argument. `as`: PptxGenJS reads `masterName` off any object.
        const masterName = typeof props === "object" && props !== null ? (props as PptxAddSlideProps).masterName : props;
        const found = this.findMaster(masterName);

        // PptxGenJS adds a slide without a master for a name it does not know, and says nothing; so does PdfLapas, but says so.
        if (found.kind === "unknown") {
            console.warn(`addSlide: no slide master is named ${JSON.stringify(found.name)}, so the slide is added without one, as PptxGenJS does. Call pres.defineSlideMaster({ title: ${JSON.stringify(found.name)}, ... }) before adding slides on it.`);
        }

        const master = found.kind === "master" ? found.master : null;
        const slideIndex = this.slides.length;
        const slide = new PuppeteerSlide({
            ...this.slideContext(slideIndex, master),
            // addTableDefinition adds autoPage continuation slides on the master of the table's slide.
            continuationSlide: offset => {
                let continuation = this.slides[slideIndex + offset];

                while (!continuation) {
                    this.addSlide(found.kind === "master" ? { masterName: found.name } : undefined);
                    continuation = this.slides[slideIndex + offset];
                }

                return continuation;
            },
            setMasterSlideNumber: props => {
                this.masterSlideNumber = props;
            },
        });

        this.slides.push(slide);

        return slide;
    }

    /** What every slide, and the layer of every master, reads from the presentation. */
    private slideContext(slideIndex: number, master: SlideMaster | null): Omit<SlideContext, "continuationSlide" | "setMasterSlideNumber"> {
        return {
            pageSize: this.pageSize,
            slideIndex,
            components: this.components,
            bodyFontFace: () => this.bodyFontFace(),
            headFontFace: () => this.themeProps?.headFontFace || PPTX_DEFAULTS.theme.headFontFace,
            warn: message => this.warningLines.add(message),
            target: this.target,
            fontMetrics: this.fontMetrics,
            sharedRunOptions: this.sharedRunOptions,
            getDocument: this.componentDocument,
            master,
        };
    }

    /**
     * The layout a master name selects, as PptxGenJS 4.0.1 finds it in `slideLayouts`: none for a falsy name, the
     * built-in "DEFAULT" (a plain slide with 0.5in table margins) before any defined master, then the first master with
     * that title.
     */
    private findMaster(name: unknown): { kind: "master"; name: string; master: SlideMaster } | { kind: "default" } | { kind: "none" } | { kind: "unknown"; name: string } {
        if (!name) {
            return { kind: "none" };
        }

        if (name === PPTX_DEFAULTS.presentation.layoutName) {
            return { kind: "default" };
        }

        const master = this.masters.get(String(name));

        return master ? { kind: "master", name: String(name), master } : { kind: "unknown", name: String(name) };
    }

    /**
     * PptxGenJS `defineLayout()`: registers the size under its name for `layout`. Unlike PptxGenJS, it also applies the
     * size at once, as PdfLapas always has; PptxGenJS code assigns `layout` next, which gives the same page.
     */
    defineLayout(layout: PptxGenJS.PresLayout): void {
        // PptxGenJS warns and writes a NaN slide size; a PDF needs a real one.
        this.pageSize = checkedPageSize("defineLayout()", layout);
        this.presLayoutName = this.pageSize.name;
        if (layout.name) {
            this.definedLayouts.set(layout.name, this.pageSize);
        }
    }

    /**
     * PptxGenJS `defineSlideMaster()`: a layout `addSlide({ masterName: title })` adds slides on. Each such slide draws
     * the master's background unless it sets its own, then the master's `objects` in order under its own, and starts
     * with the master's `slideNumber`. `margin` is where autoPage tables on those slides start and stop. A `placeholder`
     * draws nothing until a slide fills it, which throws for now (#465).
     */
    defineSlideMaster(props: PptxSlideMasterProps): void {
        // PptxGenJS deep-clones the props through JSON (ISSUE#406), so later changes to them do not reach the master.
        const clone: unknown = props === undefined ? undefined : JSON.parse(JSON.stringify(props));

        if (typeof clone !== "object" || clone === null || !("title" in clone) || !clone.title) {
            throw new Error("defineSlideMaster() object argument requires a `title` value, the name addSlide({ masterName }) selects it by. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)");
        }

        // `as`: the clone is the caller's props, which PptxGenJS reads as SlideMasterProps; each part is checked as it is read.
        const master = clone as PptxSlideMasterProps;
        // PptxGenJS: `margin || DEF_SLIDE_MARGIN_IN`.
        const margin: unknown = master.margin || PPTX_DEFAULTS.slide.marginIn;

        if (!(typeof margin === "number" || (Array.isArray(margin) && margin.length === 4 && margin.every(side => typeof side === "number")))) {
            throw new TypeError(`defineSlideMaster({ title: ${JSON.stringify(master.title)} }) \`margin\` must be inches, a number or [top, right, bottom, left], got ${JSON.stringify(margin)}`);
        }

        const layer = new PuppeteerSlide({
            ...this.slideContext(this.slides.length, null),
            continuationSlide: () => {
                throw new Error("A slide master has no autoPage continuation slides: PptxGenJS master objects hold no table.");
            },
            // Only slides write the deck's master slide number; slides added on this master copy `slideNumber` below.
            setMasterSlideNumber: () => {},
        });

        // createSlideMaster (step 1) and defineSlideMaster (step 3): the deprecated `bkgd`, then `background`.
        if (master.bkgd) {
            // `as`: the setter also takes the deprecated object, as PptxGenJS's does, which its declaration leaves out.
            layer.bkgd = master.bkgd as string;
        }

        if (master.background || master.bkgd) {
            layer.background = master.background ?? {};
        }

        const placeholders = new Map<string, MasterPlaceholder>();
        // createSlideMaster names a placeholder without `objectName` after the text objects before it, which in
        // PptxGenJS include shapes (addShapeDefinition's type is text).
        let textObjects = 0;

        // createSlideMaster (step 2) adds each object by its first key, in order; PptxGenJS skips any other key.
        for (const [index, object] of (Array.isArray(master.objects) ? master.objects : []).entries()) {
            const key = Object.keys(object)[0];

            if ("chart" in object && key === "chart") {
                // `as`: the runtime reads `{ type, data, opts }`, which the declaration's IChartOpts leaves out.
                const chart = object.chart as unknown as { type: PptxGenJS.CHART_NAME | PptxGenJS.IChartMulti[]; data: unknown[]; opts?: PptxGenJS.IChartOpts };

                layer.addChart(chart.type, chart.data, chart.opts);
            } else if ("image" in object && key === "image") {
                layer.addImage(object.image);
            } else if ("line" in object && key === "line") {
                layer.addShape("line", object.line);
                textObjects += 1;
            } else if ("rect" in object && key === "rect") {
                layer.addShape("rect", object.rect);
                textObjects += 1;
            } else if ("text" in object && key === "text") {
                // createSlideMaster writes `text` as one run without options.
                layer.addText([{ text: object.text.text ?? "" }], object.text.options);
                textObjects += 1;
            } else if ("placeholder" in object && key === "placeholder") {
                // A placeholder draws nothing: a layout's prompt text is not shown on its slides (LibreOffice export
                // checked). Text that fills it takes its options.
                const options: unknown = object.placeholder.options;

                if (typeof options !== "object" || options === null || !("name" in options) || typeof options.name !== "string") {
                    throw new TypeError(`defineSlideMaster({ title: ${JSON.stringify(master.title)} }) objects[${index}].placeholder needs \`options: { name, type, x, y, w, h }\`, the name addText(text, { placeholder: name }) fills it by.`);
                }

                // `as`: the clone of the caller's PlaceholderProps, whose `name` was checked above; PptxGenJS also reads
                // the text options its declaration leaves out.
                const { name, type, ...style } = options as PptxGenJS.PlaceholderProps & Pick<PptxGenJS.TextPropsOptions, "objectName" | "lineSpacing" | "lineSpacingMultiple">;

                if (!placeholders.has(name)) {
                    // addTextDefinition (cleanOpts) writes these onto the placeholder's options, so they replace the
                    // filling text's: no bullet and no line spacing unless the placeholder has them.
                    placeholders.set(name, {
                        title: type === "title",
                        options: {
                            ...style,
                            objectName: style.objectName ?? `Text ${textObjects}`,
                            bullet: style.bullet || false,
                            lineSpacing: style.lineSpacing || undefined,
                            lineSpacingMultiple: style.lineSpacingMultiple || undefined,
                        },
                    });
                }
            } else {
                console.warn(`defineSlideMaster: objects[${index}] has key ${JSON.stringify(key)}, which PptxGenJS ignores, so it draws nothing. Use chart, image, line, rect, text or placeholder.`);
            }
        }

        // PptxGenJS pushes every master and finds the first one with a title (`slideLayouts.filter(...)[0]`).
        // `as`: `margin` was checked above to be a number or four numbers.
        if (!this.masters.has(master.title)) {
            this.masters.set(master.title, { layer, slideNumber: master.slideNumber ?? null, margin: margin as PptxGenJS.Margin, placeholders });
        }

        // defineSlideMaster (step 4): the deck's master takes the first master's slide number while it has none.
        if (master.slideNumber && !this.masterSlideNumber) {
            this.masterSlideNumber = master.slideNumber;
        }
    }

    /**
     * PptxGenJS `tableToSlides()`: the table, paginated over as many slides as it needs. `source` is the id of a
     * <table> in the global document, as in PptxGenJS (in Node, a jsdom `document` you install),
     * the <table> element itself, or HTML markup holding one (a string starting with `<`). Markup over
     * `props.htmlLimits` (default TABLE_HTML_LIMITS) throws before it is parsed.
     */
    tableToSlides(source: string | HTMLTableElement, props?: (PptxTableToSlidesProps & { htmlLimits?: TableHtmlLimits }) | undefined): void {
        // PptxGenJS takes the master's margins (the built-in layout's are 0.5in) over `slideMargin` (genTableToSlides,
        // step 1), and adds every page on that master.
        const found = this.findMaster(props?.masterSlideName);
        const slideProps = found.kind === "none" ? undefined : { masterName: String(props?.masterSlideName) };
        const masterMargin = found.kind === "master" ? found.master.margin : PPTX_DEFAULTS.slide.marginIn;
        const options = found.kind === "master" || found.kind === "default" ? { ...props, slideMargin: masterMargin } : (props ?? {});
        const table = normalizeHtmlTable(findHtmlTable(source, this.createDocument, props?.htmlLimits), options, this.pageSize);
        const firstIndex = this.slides.length;
        const firstSlide = this.addSlide(slideProps);
        const pages = paginateTableRows(
            table.rows,
            {
                ...options,
                x: table.x,
                y: table.y,
                colW: table.columnWidths,
                autoPageRepeatHeader: options.autoPageRepeatHeader ?? options.addHeaderToEach ?? false,
                autoPageHeaderRows: table.headerRows || 1,
                autoPageSlideStartY: table.continuationY,
            },
            this.pageSize,
        );

        for (const [index, page] of pages.entries()) {
            // genTableToSlides (step 5 C) adds each page with only x, y, w and colW: `h`, `rowH` and the other table
            // options size the pages but never reach the drawn rows, which size to their content. `w` is left out
            // because a full `colW` sets the width.
            // A cell without a computed font size (jsdom) is written without `sz`; the table-level size only fills those
            // cells, after pagination counted them at PptxGenJS's 12pt.
            (index === 0 ? firstSlide : this.addSlide(slideProps)).addTable(page.rows, { x: table.x, y: page.y, colW: table.columnWidths, fontSize: PPTX_DEFAULTS.table.unsizedRunFontSizePt });
        }

        // The first slide and the continuation slides added after it.
        const generatedSlides = this.slides.slice(firstIndex);

        for (const slide of generatedSlides) {
            if (options.addImage?.image && (options.addImage.image.path || options.addImage.image.data)) {
                slide.addImage({
                    ...options.addImage.image,
                    ...(options.addImage.options ?? {}),
                });
            }

            if (options.addShape) {
                slide.addShape(options.addShape.shapeName, options.addShape.options ?? {});
            }

            if (options.addTable) {
                slide.addTable(options.addTable.rows, options.addTable.options ?? {});
            }

            if (options.addText) {
                slide.addText(options.addText.text, options.addText.options ?? {});
            }
        }
    }
}
