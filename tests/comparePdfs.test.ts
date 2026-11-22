import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtemp, readdir, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { closeBrowser } from "../packages/puppeteer/src/browser";
import { alignBaselines, comparePdfs, cubicTurns, MOVE_TOLERANCE_PT, PIXEL_THRESHOLD, readGraphics, readPages, SHAPE_TOLERANCE_PT } from "../scripts/compare-pdfs";

const OPTIONS = { moveTolerance: MOVE_TOLERANCE_PT, shapeTolerance: SHAPE_TOLERANCE_PT, pixelThreshold: PIXEL_THRESHOLD };
const TEXT = "Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod";
const ready = ["pdftotext", "pdftocairo", "pdftoppm", "magick"].every(tool => Bun.which(tool) !== null);

/** A slide with a wrapped Arial paragraph and a filled rectangle; each variant changes one thing. */
async function deck(directory: string, name: string, change: { text?: string; textWidth?: number; textY?: number; shapeX?: number; noteY?: number }): Promise<string> {
    const pres = new PuppeteerGen();
    const slide = pres.addSlide();

    slide.addText(change.text ?? TEXT, { x: 0.5, y: change.textY ?? 0.5, w: change.textWidth ?? 3, h: 1.5, fontFace: "Arial", fontSize: 14 });
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

    test("reports every word of a moved paragraph by the offset of its baseline", async () => {
        const report = await comparePdfs(await deck(directory, "base", {}), await deck(directory, "down", { textY: 0.6 }), OPTIONS);
        const page = report.pageReports[0];

        expect(page?.lineBreaks).toEqual([]);
        expect(page?.movedWords.map(word => word.word)).toEqual(TEXT.split(" "));

        for (const word of page?.movedWords ?? []) {
            expect(word.dx).toBeCloseTo(0, 1);
            // 0.1in is 9.6 CSS px; Chromium lays the text on a whole or half pixel, 7.125pt or 7.5pt lower.
            expect(word.dy).toBeCloseTo(7.2, 0);
        }
    }, 30_000);

    test("writes one deterministic diff image per page that differs, and none for a page that does not", async () => {
        const base = await deck(directory, "base", {});
        const shifted = await deck(directory, "shape", { shapeX: 5.2 });
        const diffDir = join(directory, "diffs");
        const first = await comparePdfs(base, shifted, { ...OPTIONS, diffDir });
        const image = join(diffDir, "page-1-diff.png");
        const bytes = await readFile(image);

        expect(first.pageReports[0]?.diffImage).toBe(image);
        expect(bytes.subarray(1, 4).toString()).toBe("PNG");
        await comparePdfs(base, shifted, { ...OPTIONS, diffDir });
        expect((await readFile(image)).equals(bytes)).toBe(true);

        const sameDir = join(directory, "same-diffs");
        const same = await comparePdfs(base, await deck(directory, "same", {}), { ...OPTIONS, diffDir: sameDir });

        expect(same.pageReports[0]?.diffImage).toBeUndefined();
        expect(await readdir(sameDir).catch(() => [])).toEqual([]);
    }, 60_000);
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
            { text: "ANNUAL", x: 36, y: 32.25, height: 10.035, endsLine: false },
            { text: "REPORT", x: 83.457, y: 32.25, height: 10.035, endsLine: true },
        ]);
    });
});

describe("alignBaselines", () => {
    test("puts two fonts with different descents on the same baseline", () => {
        // Arial and Times New Roman at 14pt on one baseline (y 95.6): pdftotext's box bottoms differ by 0.06pt, a
        // font with a deeper descent by more; the glyph origins cairo reads agree.
        const arial = readPages('<page width="720" height="405"><line xMin="43.2"><word xMin="43.2" yMin="82.95" xMax="82.88" yMax="98.59">Lorem</word></line></page>');
        const deep = readPages('<page width="720" height="405"><line xMin="43.2"><word xMin="43.2" yMin="80.1" xMax="82.88" yMax="99.9">Lorem</word></line></page>');
        const origins = [
            { x: 30, y: 95.6 },
            { x: 43.29, y: 95.6 },
            { x: 51.08, y: 95.6 },
            { x: 43.29, y: 120 },
        ];
        const arialWords = arial[0]?.words ?? [];
        const deepWords = deep[0]?.words ?? [];

        alignBaselines(arialWords, origins);
        alignBaselines(deepWords, origins);
        expect(arialWords[0]?.y).toBe(95.6);
        expect(deepWords[0]?.y).toBe(95.6);
    });

    test("keeps the box bottom of a word with no glyph origin inside it", () => {
        const words = readPages('<page width="720" height="405"><line xMin="43.2"><word xMin="43.2" yMin="82.95" xMax="82.88" yMax="98.59">Lorem</word></line></page>')[0]?.words ?? [];

        alignBaselines(words, [{ x: 200, y: 95.6 }]);
        expect(words[0]?.y).toBe(98.59);
    });
});

describe("readGraphics", () => {
    test("boxes each subpath, so one path of many lines matches one path per line", () => {
        const joined = '<svg><defs></defs><path d="M 10 10 L 90 10 M 10 20 L 90 20 Z M 0 0"/></svg>';
        const separate = '<svg><defs></defs><path d="M 10 10 L 90 10"/><path d="M 10 20 L 90 20"/></svg>';

        expect(readGraphics(joined, 100, 50).boxes).toEqual(readGraphics(separate, 100, 50).boxes);
        expect(readGraphics(separate, 100, 50).boxes).toHaveLength(2);
    });

    test("bounds a Bézier by the curve, not by its control points", () => {
        // An arch from (10, 30) to (50, 30) whose control points sit at y = 10: the curve peaks at y = 15.
        const svg = '<svg><defs></defs><path d="M 10 30 C 10 10 50 10 50 30 Z M 10 30"/></svg>';

        expect(readGraphics(svg, 100, 50).boxes).toEqual([{ kind: "shape", x0: 10, y0: 15, x1: 50, y1: 30 }]);
    });

    test("cubicTurns finds where a curve turns in x and in y", () => {
        // A quarter-circle-like arc turns nowhere strictly inside; the arch turns once in y, at its middle; an S curve
        // whose control points pull it past both ends turns twice in x and never in y, which runs straight.
        expect(cubicTurns({ x: 0, y: 0 }, { x: 5, y: 0 }, { x: 10, y: 5 }, { x: 10, y: 10 })).toEqual([]);
        expect(cubicTurns({ x: 10, y: 30 }, { x: 10, y: 10 }, { x: 50, y: 10 }, { x: 50, y: 30 })).toEqual([{ x: 30, y: 15 }]);

        const turns = cubicTurns({ x: 0, y: 0 }, { x: 60, y: 10 }, { x: -50, y: 20 }, { x: 10, y: 30 });

        expect(turns).toHaveLength(2);
        for (const turn of turns) {
            expect(turn.y).toBeGreaterThan(0);
            expect(turn.y).toBeLessThan(30);
        }

        expect(Math.max(turns[0]?.x ?? 0, turns[1]?.x ?? 0)).toBeGreaterThan(10);
        expect(Math.min(turns[0]?.x ?? 0, turns[1]?.x ?? 0)).toBeLessThan(0);
    });

    test("applies group and element matrices, reads image sizes and glyph origins and skips the page background", () => {
        const svg = [
            '<svg><defs><g id="glyph-0-0"><path d="M 0 0 L 5 5"/></g><image id="source-1" width="2" height="1"/></defs>',
            '<path d="M 0 0 L 100 0 L 100 50 L 0 50 Z"/>',
            '<g transform="matrix(2, 0, 0, 2, 10, 20)"><path d="M 1 1 L 3 1 L 3 2 Z" transform="matrix(1, 0, 0, 1, 1, 0)"/></g>',
            '<use xlink:href="#glyph-0-0" x="5" y="5"/>',
            '<g transform="matrix(1, 0, 0, 1, 0, 10)"><use xlink:href="#glyph-0-0" x="7" y="5"/></g>',
            '<use xlink:href="#source-1" transform="matrix(10, 0, 0, 10, 30, 30)"/></svg>',
        ].join("\n");

        expect(readGraphics(svg, 100, 50)).toEqual({
            boxes: [
                { kind: "shape", x0: 14, y0: 22, x1: 18, y1: 24 },
                { kind: "image", x0: 30, y0: 30, x1: 50, y1: 40 },
            ],
            origins: [
                { x: 5, y: 5 },
                { x: 7, y: 15 },
            ],
        });
    });
});
