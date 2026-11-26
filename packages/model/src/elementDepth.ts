// The element nesting every HTML parse of caller input stops at. parse5 is quadratic in nesting depth (100k nested
// <div> took 14 to 28 s) and the code that reads the tree walks it recursively, so deeper input is refused while it
// is parsed. The parser is passed in, so the model keeps no dependency and each package hands over its own parse5.

/** Deepest element nesting a PdfLapas HTML parse reads. `toHtml()` nests about a dozen levels. */
export const MAX_ELEMENT_DEPTH = 256;

/** The links of a parsed node the depth check follows; parse5's nodes have them. */
type TreeNode = { readonly parentNode?: TreeNode | null; readonly childNodes?: readonly TreeNode[]; readonly tagName?: string };

/** The tree adapter hooks a depth-limited parse wraps; parse5's `TreeAdapter` has them. */
type TreeHooks = {
    appendChild(parent: TreeNode, node: TreeNode): void;
    insertBefore(parent: TreeNode, node: TreeNode, reference: TreeNode): void;
    setTemplateContent(template: TreeNode, content: TreeNode): void;
};

/** The root a parse returned, or `too-deep` when its elements nest deeper than MAX_ELEMENT_DEPTH. */
export type DepthLimitedTree<Root> = { readonly kind: "tree"; readonly root: Root } | { readonly kind: "too-deep" };

/** Thrown by the wrapped adapter to stop the parser mid-parse; never leaves parseWithinDepth. */
const TOO_DEEP = Symbol("too-deep");

/**
 * Runs `parse` with `adapter` wrapped so it stops as soon as an element nests deeper than `maxDepth` below the root it
 * returns. A `fragment` parse (parse5's `parseFragment`) builds below a mock document and a fake `<html>`, one ancestor
 * more than the fragment it returns. Markup placed inside a document passes what is left of MAX_ELEMENT_DEPTH there.
 */
export function parseWithinDepth<Adapter extends TreeHooks, Root extends TreeNode>(adapter: Adapter, parse: (adapter: Adapter) => Root, kind: "document" | "fragment", maxDepth = MAX_ELEMENT_DEPTH): DepthLimitedTree<Root> {
    const limit = maxDepth + (kind === "fragment" ? 1 : 0);
    // A <template>'s content is a fragment with no parent: counting on through the template keeps nested templates
    // from restarting the count (each content counts as one more level). Local to this parse.
    const templates = new Map<TreeNode, TreeNode>();
    const checkDepth = (node: TreeNode): void => {
        // Only elements nest; text or a comment inside the deepest allowed element is fine, as in the walk below.
        if (node.tagName === undefined) {
            return;
        }

        let ancestors = 0;

        for (let parent = node.parentNode ?? templates.get(node); parent; parent = parent.parentNode ?? templates.get(parent)) {
            ancestors++;

            if (ancestors > limit) {
                throw TOO_DEEP;
            }
        }
    };
    const limited: Adapter = {
        ...adapter,
        appendChild(parent: TreeNode, node: TreeNode) {
            adapter.appendChild(parent, node);
            checkDepth(node);
        },
        insertBefore(parent: TreeNode, node: TreeNode, reference: TreeNode) {
            adapter.insertBefore(parent, node, reference);
            checkDepth(node);
        },
        setTemplateContent(template: TreeNode, content: TreeNode) {
            adapter.setTemplateContent(template, content);
            templates.set(content, template);
        },
    };
    let root: Root;

    try {
        root = parse(limited);
    } catch (error) {
        if (error === TOO_DEEP) {
            return { kind: "too-deep" };
        }

        throw error;
    }

    // The exact check, since the parser can move a subtree below a deeper parent (the adoption agency algorithm).
    // Depth first with an explicit stack: this walk must not recurse itself.
    const stack: { node: TreeNode; depth: number }[] = [{ node: root, depth: 0 }];

    for (let entry = stack.pop(); entry; entry = stack.pop()) {
        if (entry.depth > maxDepth) {
            return { kind: "too-deep" };
        }

        for (const child of entry.node.childNodes ?? []) {
            if (child.tagName !== undefined) {
                stack.push({ node: child, depth: entry.depth + 1 });
            }
        }
    }

    return { kind: "tree", root };
}
