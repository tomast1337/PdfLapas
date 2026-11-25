// The parsed profile document: parse5's tree, read once per render, and the few walks the renderer needs.
import { type DefaultTreeAdapterMap, defaultTreeAdapter, parse, parseFragment } from "parse5";

export type Element = DefaultTreeAdapterMap["element"];
type ChildNode = DefaultTreeAdapterMap["childNode"];

/**
 * Deepest element nesting read. The profile nests about a dozen levels and the renderer walks the tree recursively, so
 * deeper input is refused up front instead of exhausting the stack.
 */
export const MAX_DEPTH = 256;

/** The top-level elements of a parsed tree, or `too-deep` when elements nest deeper than MAX_DEPTH. */
export type Tree = { readonly kind: "tree"; readonly elements: Element[] } | { readonly kind: "too-deep" };

/** Thrown by the tree adapter to stop parse5 mid-parse; never leaves parseTree. */
const TOO_DEEP = Symbol("too-deep");

/** Parses a whole document (`document`) or a fragment such as an SVG image (`fragment`). */
export function parseTree(text: string, kind: "document" | "fragment"): Tree {
    // parse5 is quadratic in nesting depth (100k nested divs take seconds), so the adapter stops the parse as soon as an
    // attached node has more ancestors than allowed. A fragment parses below a mock document and a fake <html>, one
    // ancestor more than the fragment it returns. The walk below stays the exact check, since parse5 can move subtrees.
    const limit = MAX_DEPTH + (kind === "fragment" ? 1 : 0);
    const checkDepth = (node: ChildNode): void => {
        let ancestors = 0;

        for (let parent = node.parentNode; parent; parent = "parentNode" in parent ? parent.parentNode : null) {
            ancestors++;

            if (ancestors > limit) {
                throw TOO_DEEP;
            }
        }
    };
    const treeAdapter: typeof defaultTreeAdapter = {
        ...defaultTreeAdapter,
        appendChild(parent, node) {
            defaultTreeAdapter.appendChild(parent, node);
            checkDepth(node);
        },
        insertBefore(parent, node, reference) {
            defaultTreeAdapter.insertBefore(parent, node, reference);
            checkDepth(node);
        },
    };
    let parsed: DefaultTreeAdapterMap["document"] | DefaultTreeAdapterMap["documentFragment"];

    try {
        parsed = kind === "document" ? parse(text, { treeAdapter }) : parseFragment(text, { treeAdapter });
    } catch (error) {
        if (error === TOO_DEEP) {
            return { kind: "too-deep" };
        }

        throw error;
    }

    const root = parsed as unknown as Element; // `as`: a parse5 document or fragment has childNodes like an element; only they are read
    // Depth first with an explicit stack: this walk must not recurse itself.
    const stack: { element: Element; depth: number }[] = [{ element: root, depth: 0 }];

    for (let entry = stack.pop(); entry; entry = stack.pop()) {
        if (entry.depth > MAX_DEPTH) {
            return { kind: "too-deep" };
        }

        for (const child of children(entry.element)) {
            stack.push({ element: child, depth: entry.depth + 1 });
        }
    }

    return { kind: "tree", elements: children(root) };
}

export function isElement(node: ChildNode): node is Element {
    return "tagName" in node;
}

export function children(element: Element): Element[] {
    return element.childNodes.filter(isElement);
}

/** The attribute's value, or `absent` when `element` does not have it. */
export function attribute(element: Element, name: string, absent: string): string {
    return element.attrs.find(attr => attr.name === name)?.value ?? absent;
}

export function hasAttribute(element: Element, name: string): boolean {
    return element.attrs.some(attr => attr.name === name);
}

export function hasClass(element: Element, name: string): boolean {
    return attribute(element, "class", "").split(/\s+/).includes(name);
}

/** Every element below `roots` in document order, depth first. */
export function* descendants(roots: readonly Element[]): Generator<Element> {
    for (const root of roots) {
        yield root;
        yield* descendants(children(root));
    }
}

/** The concatenated text of `element`'s text nodes, its descendants' included. */
export function textContent(element: Element): string {
    return element.childNodes.map(node => (isElement(node) ? textContent(node) : "value" in node ? node.value : "")).join("");
}
