import type PptxGenJS from "pptxgenjs";
import type { ComponentRegistry } from "./components/registry";
import type { FontMetricsLookup } from "./defaults";
import type { RenderingTarget } from "./target";
import type { SharedRunOptions } from "./normalize/sharedRunOptions";
import type { PageSize } from "./pageLayouts";
import type { PuppeteerSlide } from "./PuppeteerSlide";

/** A defineSlideMaster() layout: PptxGenJS draws its background and objects under those of each slide added on it. */
export type SlideMaster = {
    /** The master's background and objects, normalized as a slide of their own that is never a page of the deck; null for PptxGenJS's built-in "DEFAULT", which has neither. */
    readonly layer: PuppeteerSlide | null;
    /** The `slideNumber` each slide added on the master starts with. */
    readonly slideNumber: PptxGenJS.SlideNumberProps | null;
    /** `margin` in inches, [top, right, bottom, left] or one for all: where autoPage tables on its slides start and stop. */
    readonly margin: PptxGenJS.Margin;
    /** The master's `placeholder` objects by `name`, the first one under a name, as addText({ placeholder }) finds it. */
    readonly placeholders: ReadonlyMap<string, MasterPlaceholder>;
};

/** A slide master's `placeholder` object, which text filling it takes its options from. */
export type MasterPlaceholder = {
    /** Whether its `type` is "title", which inherits the master's titleStyle; every other type the bodyStyle. */
    readonly title: boolean;
    /** Its `options` without `name` and `type`, as PptxGenJS's createSlideMaster and addTextDefinition clean them. */
    readonly options: PptxGenJS.TextPropsOptions;
    /**
     * PptxGenJS 4.0.1's `_bodyProp` of the placeholder, which addTextDefinition (A.3) shares by reference with every
     * text that fills it, on any slide: each addText() writes `anchor`, `vert`, `wrap` and `autoFit` onto it (`inset`
     * only when given), so the last fill's are what every fill's <a:bodyPr> says when the file is written. Mutable,
     * as that object is; the presentation owns it with the master, and each fill replaces it.
     */
    body: PlaceholderBody;
};

/** What a filled placeholder's shared `_bodyProp` holds when the file is written, as the filling text's options. */
export type PlaceholderBody = Pick<PptxGenJS.TextPropsOptions, "valign" | "vert" | "wrap" | "autoFit" | "inset">;

/** A fill's `margin` as slideObjectToXml writes it onto the shared `_bodyProp`; undefined: none written yet. */
export type PlaceholderMargin = PptxGenJS.TextPropsOptions["margin"];

/** What a slide reads from the presentation that owns it. PuppeteerGen.addSlide() creates one per slide; every field is required. */
export type SlideContext = {
    /** The presentation's page size when the slide was added, which its percentage positions resolve against. */
    readonly pageSize: PageSize;
    /** The slide's 0-based position: its slide number and the scope of its object ids. */
    readonly slideIndex: number;
    /** The presentation's registered components, which addComponent() renders once to check them. */
    readonly components: ComponentRegistry;
    /** `pres.theme.bodyFontFace`: the font of text, table cells and chart legends that name none (PowerPoint's `+mn-lt`). */
    readonly bodyFontFace: () => string;
    /** `pres.theme.headFontFace`: the font of title placeholder text that names none (PowerPoint's `+mj-lt`). */
    readonly headFontFace: () => string;
    /** Adds a line to the presentation's warnings(): what PptxGenJS ignores without a word. */
    readonly warn: (message: string) => void;
    /** The presentation's rendering target, which text is normalized for. */
    readonly target: RenderingTarget;
    /** The faces' metrics (`fontMetrics`, `resolveFontMetrics`, FONT_WIN_METRICS) the target may place baselines from. */
    readonly fontMetrics: FontMetricsLookup;
    /** What PptxGenJS has written onto the run options objects of the presentation's text boxes and table cells. */
    readonly sharedRunOptions: SharedRunOptions;
    /** The `shadow` objects the presentation's text boxes, shapes and pictures were given so far. */
    readonly shadowsSeen: WeakSet<object>;
    /**
     * The chart options objects the presentation's addChart() calls were given. PptxGenJS 4.0.1 keeps a reference to
     * each one and reads it only when it writes the file, so a second chart on one object warns (PuppeteerSlide.addChart).
     */
    readonly chartOptionsSeen: WeakSet<object>;
    /** The `context.document` of components rendered to check them when they are added; toHtml() does not use it. */
    readonly getDocument: () => Document;
    /** The slide `offset` places after this one, added when it does not exist yet: where autoPage tables continue. */
    readonly continuationSlide: (offset: number) => PuppeteerSlide;
    /** PptxGenJS's `setSlideNumber` callback: every slide's `slideNumber` setter also sets the master's. */
    readonly setMasterSlideNumber: (props: PptxGenJS.SlideNumberProps | null) => void;
    /** The defineSlideMaster() layout the slide was added on (`addSlide({ masterName })`), or null for none. */
    readonly master: SlideMaster | null;
};
