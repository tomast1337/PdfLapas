import { type DataUri, MAX_ELEMENT_DEPTH } from "@pdflapas/model";
import type { DefaultTreeAdapterTypes } from "parse5";
import type { normalizeComponent } from "#src/components/normalize";
import type { ComponentDefinition } from "#src/components/types";
import { COMPONENT_DEFAULTS } from "#src/defaults";
import { type InlineStyle, styleAttribute } from "./css";
import { descendants, type FragmentElement, getAttribute, HTML_NAMESPACE, parseMarkup, removeAttribute, serialize, setAttribute } from "./fragment";
import { type Attribute, element, type Html } from "./markup";
import { applyObjectStyle } from "./style";

const REMOTE_URL = /^https?:\/\//i;
// The elements above a component's content in toHtml(): <html>, <body>, the .slide-container and the component's host.
// Its markup gets what is left of MAX_ELEMENT_DEPTH, so the deck's HTML stays within what every parser reads.
const COMPONENT_CONTENT_ANCESTORS = 4;

/** A rendered component: its host attributes and its parsed content, whose images resolve before it is written. */
export type RenderedComponent = {
    attributes: Attribute[];
    content: DefaultTreeAdapterTypes.DocumentFragment;
};

function images(content: DefaultTreeAdapterTypes.DocumentFragment): FragmentElement[] {
    return Array.from(descendants(content)).filter(node => node.tagName === "img" && node.namespaceURI === HTML_NAMESPACE);
}

/**
 * `idScope` is the object's position in the deck (`objectIdScope`), so ids the component derives from it are
 * unique and stable. `getDocument` is called only when the component reads `context.document`, so a component
 * that returns HTML text needs no DOM.
 */
export function renderComponent(component: ReturnType<typeof normalizeComponent>, definition: ComponentDefinition, idScope: string, getDocument: () => Document): RenderedComponent {
    const { input, style } = component;
    // A copy, so a definition that changes its props cannot change the model: toHtml() renders the presentation's own.
    const props = structuredClone(input.props);

    definition.validate?.(props);
    const context = {
        get document() {
            return getDocument();
        },
        width: style.geometry.width!,
        height: style.geometry.height!,
        idPrefix: idScope,
    };
    const result = definition.render(props, context);
    const css: InlineStyle = new Map();
    const objectAttributes = applyObjectStyle(css, style);

    // Explicit dimensions define the component viewport, independently of inner HTML sizing.
    css.set("position", "absolute");
    css.set("overflow", COMPONENT_DEFAULTS.overflow);
    const attributes: Attribute[] = [["class", "slide-element slide-component"], ["data-component-type", input.type], ["data-component-version", String(input.version)], ...objectAttributes, styleAttribute(css)];

    let markup: string;

    // Not `instanceof` for an element: browser documents from createHTMLDocument() have no defaultView to supply the
    // HTMLElement class.
    if (typeof result === "string") {
        markup = result;
    } else if ((result as Node | undefined)?.ownerDocument && result.ownerDocument === getDocument() && result.namespaceURI === HTML_NAMESPACE) {
        markup = result.outerHTML;
    } else {
        throw new Error("Component render must return HTML text or an HTMLElement from the supplied document");
    }

    // Parsed as the host's innerHTML, so the markup is normalized as a DOM would have it.
    const maxDepth = MAX_ELEMENT_DEPTH - COMPONENT_CONTENT_ANCESTORS;
    const parsed = parseMarkup(markup, "div", maxDepth);

    if (parsed.kind === "too-deep") {
        throw new Error(
            `Component '${input.type}' version ${input.version} rendered markup whose elements nest deeper than ${maxDepth} levels (the deck's HTML allows ${MAX_ELEMENT_DEPTH}, and the slide takes ${COMPONENT_CONTENT_ANCESTORS}); flatten its markup`,
        );
    }

    const content = parsed.root;

    for (const image of images(content)) {
        if (getAttribute(image, "srcset").kind === "value") {
            throw new Error("Component images do not support srcset; use a single src");
        }

        const src = getAttribute(image, "src");
        const source = src.kind === "value" ? src.value : "";

        if (source && !source.startsWith("data:")) {
            setAttribute(image, REMOTE_URL.test(source) ? "data-source-url" : "data-source-path", source);
            removeAttribute(image, "src");
        }
    }

    return { attributes, content };
}

/** The image source a component still has to resolve: `data-source-path`, then `data-source-url`; "none" when empty. */
function pendingSource(image: FragmentElement): { kind: "pending"; source: string } | { kind: "none" } {
    const path = getAttribute(image, "data-source-path");
    // A path, even an empty one, wins over a URL.
    const source = path.kind === "value" ? path : getAttribute(image, "data-source-url");

    return source.kind === "value" && source.value ? { kind: "pending", source: source.value } : { kind: "none" };
}

export function componentImageSources(component: RenderedComponent): string[] {
    const sources: string[] = [];

    for (const image of images(component.content)) {
        const pending = pendingSource(image);

        if (pending.kind === "pending") {
            sources.push(pending.source);
        }
    }

    return sources;
}

/** The host with its content; `resolved` maps image sources to data URIs, which replace them as `src`. */
export function componentHtml(component: RenderedComponent, resolved?: ReadonlyMap<string, DataUri>): Html {
    for (const image of images(component.content)) {
        const pending = pendingSource(image);
        const data = pending.kind === "pending" && resolved?.get(pending.source);

        if (!data) {
            continue;
        }

        setAttribute(image, "src", data);
        removeAttribute(image, "data-source-path");
        removeAttribute(image, "data-source-url");
    }

    return element("div", component.attributes, serialize(component.content));
}
