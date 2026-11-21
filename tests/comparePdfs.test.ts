import { afterAll, beforeAll, describe, expect, test } from "bun:test";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { closeBrowser } from "../packages/puppeteer/src/browser";
import { comparePdfs, MOVE_TOLERANCE_PT, PIXEL_THRESHOLD, readBoxes, SHAPE_TOLERANCE_PT } from "../scripts/compare-pdfs";

const OPTIONS = { moveTolerance: MOVE_TOLERANCE_PT, shapeTolerance: SHAPE_TOLERANCE_PT, pixelThreshold: PIXEL_THRESHOLD };
const TEXT = "Lorem ipsum dolor sit amet consectetur adipiscing elit sed do eiusmod";
const ready = ["pdftotext", "pdftocairo", "pdftoppm", "magick"].every(tool => Bun.which(tool) !== null);

/** A slide with a wrapped Arial paragraph and a filled rectangle; each variant changes one thing. */
async function deck(directory: string, name: string, change: { text?: string; textWidth?: number; shapeX?: number }): Promise<string> {
    const pres = new PuppeteerGen();
    const slide = pres.addSlide();

    slide.addText(change.text ?? TEXT, { x: 0.5, y: 0.5, w: change.textWidth ?? 3, h: 1.5, fontFace: "Arial", fontSize: 14 });
    slide.addShape("rect", { x: change.shapeX ?? 5, y: 1, w: 2, h: 1, fill: { color: "FF0000" } });
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

describe("readBoxes", () => {
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
