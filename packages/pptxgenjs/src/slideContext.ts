import type PptxGenJS from "pptxgenjs";
import type { ComponentRegistry } from "./components/registry";
import type { FontMetricsLookup } from "./defaults";
import type { RenderingTarget } from "./target";
import type { SharedRunOptions } from "./normalize/sharedRunOptions";
import type { PageSize } from "./pageLayouts";
import type { PuppeteerSlide } from "./PuppeteerSlide";

/** A defineSlideMaster() layout: PptxGenJS draws its background and objects under those of each slide added on it. */
export type SlideMaster = {
    /** The master's background and objects, normalized as a slide of their own that is never a page of the deck. */
    readonly layer: PuppeteerSlide;
    /** The `slideNumber` each slide added on the master starts with. */
    readonly slideNumber: PptxGenJS.SlideNumberProps | null;
    /** `margin` in inches, [top, right, bottom, left] or one for all: where autoPage tables on its slides start and stop. */
    readonly margin: PptxGenJS.Margin;
};

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
    /** The presentation's rendering target, which text is normalized for. */
    readonly target: RenderingTarget;
    /** The faces' metrics (`fontMetrics`, `resolveFontMetrics`, FONT_WIN_METRICS) the target may place baselines from. */
    readonly fontMetrics: FontMetricsLookup;
    /** What PptxGenJS has written onto the run options objects of the presentation's text boxes and table cells. */
    readonly sharedRunOptions: SharedRunOptions;
    /** The `context.document` of components rendered to check them when they are added; toHtml() does not use it. */
    readonly getDocument: () => Document;
    /** The slide `offset` places after this one, added when it does not exist yet: where autoPage tables continue. */
    readonly continuationSlide: (offset: number) => PuppeteerSlide;
    /** PptxGenJS's `setSlideNumber` callback: every slide's `slideNumber` setter also sets the master's. */
    readonly setMasterSlideNumber: (props: PptxGenJS.SlideNumberProps | null) => void;
    /** The defineSlideMaster() layout the slide was added on (`addSlide({ masterName })`), or null for none. */
    readonly master: SlideMaster | null;
};
