// Property-based tests for PdfkitRenderer.render() (#661, fourth piece): profile HTML from toHtml(), mutated as a
// caller who builds or edits the document could (numbers, attribute and style values, cut or repeated markup, stray
// tags and characters), either renders to a PDF or is refused with an error that says what is wrong. A
// JavaScript-engine error means a value reached code that did not check it, the kind of bug that ended the process in
// #659. As crates/fuzz/render_html does for pdf-core, without coverage guidance.
import { beforeAll, describe, expect, test } from "bun:test";
import fc from "fast-check";
import type { HtmlPresentationDocument } from "../packages/model/src/index";
import type { PdfkitRenderer } from "../packages/pdfkit/src/browser";
import { createPdfkitRenderer, loadSystemFonts } from "../packages/pdfkit/src/node";
import { PuppeteerGen } from "../packages/pptxgenjs/src/html";
import { FUZZ_SEED, fuzzRuns, fuzzTimeout } from "./fuzzConfig";

const RUNS = fuzzRuns(150);

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";
const SVG = `data:image/svg+xml;base64,${Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="red"/></svg>').toString("base64")}`;

/** Messages the JavaScript engine raises for a value it did not expect: an internal failure, not a documented refusal. */
const INTERNAL_FAILURE = [
    /Cannot read propert/,
    /Cannot set propert/,
    /is not a function/,
    /is not iterable/,
    /is not an object/,
    /Cannot convert undefined or null/,
    /Cannot destructure/,
    /Cannot use 'in' operator/,
    /Invalid array length/,
    /Invalid string length/,
    /Maximum call stack size exceeded/,
    /out of memory/i,
    /Out of memory/,
    /Cannot access .* before initialization/,
    /Reduce of empty array/,
    /RangeError: Array buffer allocation failed/,
    /toFixed\(\) argument/,
    /Invalid typed array length/,
    // PDFKit's own refusal of a number past 1e21, which names no input: render() must say what in the document is wrong.
    /^unsupported number/,
];

/** Small decks whose HTML holds each element the profile writes: rich text with a link, shapes, images, a table, a chart. */
async function baseDocuments(): Promise<HtmlPresentationDocument[]> {
    const deck = new PuppeteerGen();

    deck.addSlide().addText(
        [
            { text: "Hello ", options: { bold: true, color: "FF0000", underline: { style: "dash" } } },
            { text: "world 世界 😀", options: { italic: true, hyperlink: { url: "https://example.com", tooltip: "Example" }, highlight: "FFFF00" } },
            { text: "Second paragraph", options: { breakLine: true, bullet: true, strike: "sngStrike" } },
        ],
        { x: 1, y: 1, w: 4, h: 1.5, fill: { color: "EEEEEE" }, align: "center", valign: "middle", shadow: { type: "outer", blur: 3, offset: 2, angle: 45, color: "000000", opacity: 0.4 }, fit: "shrink" },
    );
    deck.addSlide().addShape("roundRect", { x: 1, y: 1, w: 2, h: 1, fill: { color: "0088CC", transparency: 30 }, rectRadius: 0.2, rotate: 15, line: { color: "333333", width: 2, dashType: "dash" } });
    deck.addSlide().addShape("line", { x: 1, y: 1, w: 3, h: 0, line: { color: "FF0000", width: 1, beginArrowType: "triangle", endArrowType: "oval" } });
    deck.addSlide().addShape("cloud", { x: 1, y: 1, w: 3, h: 2, fill: { color: "00AA00" } });
    deck.addSlide().addImage({ data: PNG, x: 1, y: 1, w: 1, h: 1, transparency: 20, rounding: true });
    deck.addSlide().addImage({ data: SVG, x: 1, y: 1, w: 1, h: 1, sizing: { type: "crop", w: 0.5, h: 0.5 } });
    deck.addSlide().addTable(
        [
            [
                { text: "A", options: { bold: true, rowspan: 2 } },
                { text: "B", options: { colspan: 2 } },
            ],
            [{ text: "C" }, { text: "D", options: { fill: { color: "DDDDDD" }, valign: "bottom" } }],
        ],
        { x: 1, y: 1, w: 4, colW: [1, 1.5, 1.5], border: { type: "solid", pt: 1, color: "000000" }, fontSize: 12 },
    );
    deck.addSlide().addChart("bar", [{ name: "Series", labels: ["A", "B", "C"], values: [3, 5, 4] }], { x: 1, y: 1, w: 4, h: 3, showLegend: true, showValue: true });

    return [await deck.toHtml()];
}

// --- Mutating the HTML ------------------------------------------------------------------------------------------

/** Where a mutation can land: each number, each quoted attribute value, each style declaration's value, each tag. */
type Sites = { readonly numbers: readonly [number, number][]; readonly attributes: readonly [number, number][]; readonly declarations: readonly [number, number][]; readonly tags: readonly [number, number][] };

// `d`: each match carries the offsets of its groups.
const NUMBER = /(-?\d+(?:\.\d+)?(?:e[+-]?\d+)?)/dg;
const ATTRIBUTE = /="([^"]*)"/dg;
const DECLARATION = /:\s*([^;"]+)/dg;
const TAG = /(<[^<>]+>)/dg;

function sites(html: string): Sites {
    const ranges = (pattern: RegExp): [number, number][] => {
        const found: [number, number][] = [];

        for (const match of html.matchAll(pattern)) {
            const range = match.indices?.[1];

            if (range !== undefined) {
                found.push(range);
            }
        }

        return found;
    };

    return { numbers: ranges(NUMBER), attributes: ranges(ATTRIBUTE), declarations: ranges(DECLARATION), tags: ranges(TAG) };
}

type Mutation =
    | { readonly kind: "number" | "attribute" | "declaration"; readonly at: number; readonly value: string }
    | { readonly kind: "cut" | "repeat"; readonly at: number; readonly length: number; readonly times: number }
    | { readonly kind: "insert"; readonly at: number; readonly text: string }
    | { readonly kind: "char"; readonly at: number; readonly char: string };

const EXTREME_NUMBERS = ["0", "-0", "-1", "1e308", "-1e308", "1e-308", "NaN", "Infinity", "999999999", "4294967296", "0.0000001", "", "1.5.5", "1e", "-"];
const STRANGE_VALUES = [
    ...EXTREME_NUMBERS,
    "x",
    "none",
    "inherit",
    "auto",
    "calc(100% - 1px)",
    "url(x)",
    "rgb(1,2)",
    "rgba(255,0,0,2)",
    "#",
    "#12",
    "999999999px",
    "-999999999px",
    "1e308px",
    "NaNpx",
    "translate(1e308px) rotate(1e308deg)",
    "matrix(0,0,0,0,0,0)",
    "scale(0)",
    "data:image/png;base64,",
    "data:image/png;base64,AAAA",
    "data:image/svg+xml,<svg/>",
    "javascript:alert(1)",
    "a".repeat(5000),
    "😀",
    "\u0000",
    "&amp;&#xFFFF;&#0;",
];
const STRAY_TAGS = [
    "<div>",
    "</div>",
    "<span>",
    "</p>",
    "<p>",
    "<table><tr><td>x</td></tr></table>",
    "<img>",
    '<img src="data:image/png;base64,AAAA">',
    '<svg><path d="M 0 0 L 1e308 1e308"/></svg>',
    "<svg>",
    "<br>",
    '<div style="position:absolute;left:1e308px;width:-5px">x</div>',
    '<span style="font-size:0">x</span>',
    '<span style="font-size:1e9px">x</span>',
    "<!--",
    "<![CDATA[",
    "<script>x</script>",
    "<style>*{}</style>",
];

const value = fc.oneof(fc.constantFrom(...STRANGE_VALUES), fc.string({ maxLength: 12 }), fc.double().map(String), fc.integer().map(String));
const mutation: fc.Arbitrary<Mutation> = fc.oneof(
    { weight: 4, arbitrary: fc.record({ kind: fc.constant("number" as const), at: fc.nat(), value }) },
    { weight: 3, arbitrary: fc.record({ kind: fc.constant("attribute" as const), at: fc.nat(), value }) },
    { weight: 3, arbitrary: fc.record({ kind: fc.constant("declaration" as const), at: fc.nat(), value }) },
    { weight: 2, arbitrary: fc.record({ kind: fc.constant("cut" as const), at: fc.nat(), length: fc.integer({ min: 1, max: 400 }), times: fc.constant(0) }) },
    // Repeated markup nests or multiplies elements, as a runaway producer loop would.
    { weight: 2, arbitrary: fc.record({ kind: fc.constant("repeat" as const), at: fc.nat(), length: fc.integer({ min: 1, max: 400 }), times: fc.integer({ min: 2, max: 60 }) }) },
    { weight: 2, arbitrary: fc.record({ kind: fc.constant("insert" as const), at: fc.nat(), text: fc.oneof(fc.constantFrom(...STRAY_TAGS), fc.string({ maxLength: 8 })) }) },
    { weight: 1, arbitrary: fc.record({ kind: fc.constant("char" as const), at: fc.nat(), char: fc.constantFrom("<", ">", '"', "'", "&", ";", ":", "=", "/", " ", "\u0000", "￿") }) },
);
const mutations = fc.array(mutation, { minLength: 1, maxLength: 4 });

/** The HTML with each mutation applied in order, each at a site of the HTML as the previous one left it. */
function mutate(html: string, operations: readonly Mutation[]): string {
    let result = html;

    for (const operation of operations) {
        const found = sites(result);

        if (operation.kind === "number" || operation.kind === "attribute" || operation.kind === "declaration") {
            const list = operation.kind === "number" ? found.numbers : operation.kind === "attribute" ? found.attributes : found.declarations;
            const range = list[operation.at % Math.max(1, list.length)];

            if (range !== undefined) {
                result = `${result.slice(0, range[0])}${operation.value.replaceAll('"', "&quot;")}${result.slice(range[1])}`;
            }
        } else if (operation.kind === "cut" || operation.kind === "repeat") {
            // From a tag's start, so a cut or a repeat takes whole elements more often than a byte offset would.
            const range = found.tags[operation.at % Math.max(1, found.tags.length)];
            const start = range?.[0] ?? 0;
            const piece = result.slice(start, start + operation.length);

            result = `${result.slice(0, start)}${operation.kind === "repeat" ? piece.repeat(operation.times) : ""}${result.slice(start + operation.length)}`;
        } else {
            const range = found.tags[operation.at % Math.max(1, found.tags.length)];
            const at = range?.[operation.at % 2] ?? 0;

            result = `${result.slice(0, at)}${operation.kind === "insert" ? operation.text : operation.kind === "char" ? operation.char : ""}${result.slice(at)}`;
        }
    }

    return result;
}

/** Fails when `error` is a JavaScript-engine error (a crash); a documented refusal or a renderer report is accepted. */
function expectRefusal(error: unknown): void {
    if (!(error instanceof Error)) {
        throw new Error(`render() threw a value that is not an Error: ${String(error)}`);
    }

    for (const pattern of INTERNAL_FAILURE) {
        if (pattern.test(error.message) || pattern.test(String(error))) {
            throw error;
        }
    }
}

describe("PdfkitRenderer.render() over mutated profile HTML (#661)", () => {
    let renderer: PdfkitRenderer;
    let bases: HtmlPresentationDocument[];

    beforeAll(async () => {
        renderer = await createPdfkitRenderer({ fonts: await loadSystemFonts(), onUnsupported: "warn" });
        bases = await baseDocuments();
    });

    test("the base documents render", async () => {
        for (const base of bases) {
            const pdf = await renderer.render(base);

            expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe("%PDF-");
        }
    });

    test("a length a PDF cannot hold is refused as invalid HTML, not with PDFKit's message", async () => {
        const [base] = bases;
        const html = base?.html.replace(/text-decoration-thickness: [0-9.]+em/, "text-decoration-thickness: 1e20em");

        expect(html).not.toBe(base?.html);
        await expect(base && html ? renderer.render({ ...base, html }) : Promise.reject(new Error("no base document"))).rejects.toThrow(/^invalid PdfLapas HTML: a length, or one computed from it, is -?\d.*e\+\d+, which a PDF cannot hold/);
    });

    test(
        "mutated HTML renders to a PDF or is refused with an error that is not an engine error",
        async () => {
            await fc.assert(
                fc.asyncProperty(fc.nat(), mutations, async (which, operations) => {
                    const base = bases[which % bases.length];

                    if (base === undefined) {
                        return;
                    }

                    const document = { ...base, html: mutate(base.html, operations) };

                    try {
                        const pdf = await renderer.render(document);

                        expect(new TextDecoder().decode(pdf.subarray(0, 5))).toBe("%PDF-");
                    } catch (error) {
                        expectRefusal(error);
                    }
                }),
                { seed: FUZZ_SEED, numRuns: RUNS },
            );
        },
        fuzzTimeout(120_000),
    );
});
