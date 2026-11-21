import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { closeBrowser } from "../packages/puppeteer/src/browser";
import { comparePdfs, MOVE_TOLERANCE_PT, PIXEL_THRESHOLD, readBoxes, readPages, SHAPE_TOLERANCE_PT } from "../scripts/compare-pdfs";

const OPTIONS = { moveTolerance: MOVE_TOLERANCE_PT, shapeTolerance: SHAPE_TOLERANCE_PT, pixelThreshold: PIXEL_THRESHOLD };
const TEXT = "Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod";
const ready = ["pdftotext", "pdftocairo", "pdftoppm", "magick"].every(tool => Bun.which(tool) !== null);

/** A slide with a wrapped Arial paragraph and a filled rectangle; each variant changes one thing. */
async function deck(directory: string, name: string, change: { text?: string; textWidth?: number; shapeX?: number; noteY?: number }): Promise<string> {
    const pres = new PuppeteerGen();
    const slide = pres.addSlide();

    slide.addText(change.text ?? TEXT, { x: 0.5, y: 0.5, w: change.textWidth ?? 3, h: 1.5, fontFace: "Arial", fontSize: 14 });
    slide.addShape("rect", { x: change.shapeX ?? 5, y: 1, w: 2, h: 1, fill: { color: "FF0000" } });
    slide.addText("Placeholder note", { x: 0.5, y: change.noteY ?? 4.5, w: 3, h: 0.4, fontFace: "Arial", fontSize: 12 });
    const file = join(directory, `${name}.pdf`);

    await writeFile(file, await pres.toPdf());

    return file;
}

describe.skipIf(!ready)("compare-pdfs", () => {
    let directory = "";

    beforeAll(async () => {
        directory = await mkdtemp(join(tmpdir(), "compare-pdfs-"));
    });

    afterAll(async () => {
        await closeBrowser();
        await rm(directory, { recursive: true, force: true });
    });

    test("finds no difference between two renders of the same deck", async () => {
        const report = await comparePdfs(await deck(directory, "base", {}), await deck(directory, "same", {}), OPTIONS);

        expect(report.passed).toBe(true);
        expect(report.pageReports[0]).toMatchObject({ removedWords: [], addedWords: [], lineBreaks: [], movedWords: [], movedBoxes: [], missingBoxes: [], extraBoxes: [] });
    }, 30_000);

    test("reports a changed word", async () => {
        const report = await comparePdfs(await deck(directory, "base", {}), await deck(directory, "word", { text: TEXT.replace("eiusmod", "tempor") }), OPTIONS);

        expect(report.passed).toBe(false);
        expect(report.pageReports[0]).toMatchObject({ removedWords: ["eiusmod"], addedWords: ["tempor"], lineBreaks: [], movedBoxes: [] });
    }, 30_000);

    test("reports a line that breaks after another word", async () => {
        const report = await comparePdfs(await deck(directory, "base", {}), await deck(directory, "wrap", { textWidth: 2.6 }), OPTIONS);
        const page = report.pageReports[0];

        expect(report.passed).toBe(false);
        expect(page?.removedWords).toEqual([]);
        expect(page?.lineBreaks).toContainEqual({ word: "sed", endsLineIn: "reference" });
        expect(page?.movedWords.length).toBeGreaterThan(0);
    }, 30_000);

    test("matches words that another reading order puts elsewhere and reports them as moved", async () => {
        // The note moves above the paragraph, so pdftotext reads it first in one PDF and last in the other.
        const report = await comparePdfs(await deck(directory, "base", {}), await deck(directory, "note", { noteY: 0.05 }), OPTIONS);
        const page = report.pageReports[0];

        expect(page?.removedWords).toEqual([]);
        expect(page?.addedWords).toEqual([]);
        expect(page?.movedWords.map(word => word.word)).toEqual(["Placeholder", "note"]);
    }, 30_000);

    test("reports a moved shape with its offset", async () => {
        const report = await comparePdfs(await deck(directory, "base", {}), await deck(directory, "shape", { shapeX: 5.2 }), OPTIONS);
        const page = report.pageReports[0];

        expect(report.passed).toBe(false);
        expect(page?.removedWords).toEqual([]);
        expect(page?.movedWords).toEqual([]);
        expect(page?.missingBoxes).toEqual([]);
        expect(page?.movedBoxes).toHaveLength(1);
        expect(page?.movedBoxes[0]?.dx).toBeCloseTo(14.4, 0);
        expect(page?.movedBoxes[0]?.dy).toBeCloseTo(0, 0);
    }, 30_000);
});

describe("readPages", () => {
    test("joins the fragments of a letter-spaced word and keeps the word breaks", () => {
        // pdftotext's words for PowerPoint's export of "ANNUAL REPORT" with charSpacing: each glyph drawn apart.
        const fragments: [number, number, string][] = [
            [36, 49.806, "AN"],
            [50.742, 57.24, "N"],
            [58.176, 64.674, "U"],
            [65.61, 78.417, "AL"],
            [83.457, 89.955, "R"],
            [90.891, 96.894, "E"],
            [97.839, 103.842, "P"],
            [104.787, 111.789, "O"],
            [112.698, 119.196, "R"],
            [120.132, 125.631, "T"],
        ];
        const words = fragments.map(([x0, x1, text]) => `<word xMin="${x0}" yMin="22.215" xMax="${x1}" yMax="32.25">${text}</word>`).join("");
        const pages = readPages(`<page width="720.000000" height="405.000000"><flow><block><line xMin="36">${words}</line></block></flow></page>`);

        expect(pages[0]?.words).toEqual([
            { text: "ANNUAL", x: 36, y: 32.25, endsLine: false },
            { text: "REPORT", x: 83.457, y: 32.25, endsLine: true },
        ]);
    });
});

describe("readBoxes", () => {
    test("boxes each subpath, so one path of many lines matches one path per line", () => {
        const joined = '<svg><defs></defs><path d="M 10 10 L 90 10 M 10 20 L 90 20 Z M 0 0"/></svg>';
        const separate = '<svg><defs></defs><path d="M 10 10 L 90 10"/><path d="M 10 20 L 90 20"/></svg>';

        expect(readBoxes(joined, 100, 50)).toEqual(readBoxes(separate, 100, 50));
        expect(readBoxes(separate, 100, 50)).toHaveLength(2);
    });

    test("applies group and element matrices, reads image sizes and skips glyphs and the page background", () => {
        const svg = [
            '<svg><defs><g id="glyph-0-0"><path d="M 0 0 L 5 5"/></g><image id="source-1" width="2" height="1"/></defs>',
            '<path d="M 0 0 L 100 0 L 100 50 L 0 50 Z"/>',
            '<g transform="matrix(2, 0, 0, 2, 10, 20)"><path d="M 1 1 L 3 1 L 3 2 Z" transform="matrix(1, 0, 0, 1, 1, 0)"/></g>',
            '<use xlink:href="#glyph-0-0" x="5" y="5"/>',
            '<use xlink:href="#source-1" transform="matrix(10, 0, 0, 10, 30, 30)"/></svg>',
        ].join("\n");

        expect(readBoxes(svg, 100, 50)).toEqual([
            { kind: "shape", x0: 14, y0: 22, x1: 18, y1: 24 },
            { kind: "image", x0: 30, y0: 30, x1: 50, y1: 40 },
        ]);
    });
});
