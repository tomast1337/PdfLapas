// The parsed profile document: parse5's tree, read once per render, and the few walks the renderer needs.
import { type DefaultTreeAdapterMap, parse, parseFragment } from "parse5";

export type Element = DefaultTreeAdapterMap["element"];
type ChildNode = DefaultTreeAdapterMap["childNode"];

/**
 * Deepest element nesting read. The profile nests about a dozen levels and the renderer walks the tree recursively, so
 * deeper input is refused up front instead of exhausting the stack.
 */
export const MAX_DEPTH = 256;

/** The top-level elements of a parsed tree, or `too-deep` when elements nest deeper than MAX_DEPTH. */
export type Tree = { readonly kind: "tree"; readonly elements: Element[] } | { readonly kind: "too-deep" };

/** Parses a whole document (`document`) or a fragment such as an SVG image (`fragment`). */
export function parseTree(text: string, kind: "document" | "fragment"): Tree {
    const root = (kind === "document" ? parse(text) : parseFragment(text)) as unknown as Element; // `as`: a parse5 document or fragment has childNodes like an element; only they are read
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
