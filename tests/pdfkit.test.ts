import { describe, expect, test } from "bun:test";
import { crc32, deflateSync, inflateSync } from "node:zlib";
import { create as createFont, type Font } from "fontkit";
import { JSDOM } from "jsdom";
import PDFDocument from "pdfkit";
import type PptxGenJS from "pptxgenjs";
import { PuppeteerGen } from "../packages/pptxgenjs/src/html";
import { PuppeteerGen as NodePuppeteerGen } from "../packages/pptxgenjs/src/node";
import { PAGE_SIZES } from "../packages/pptxgenjs/src/pageLayouts";
import { createPdfkitRenderer as createBrowserPdfkitRenderer, loadSystemFonts as browserLoadSystemFonts, type PdfkitRenderer } from "../packages/pdfkit/src/browser";
import { createPdfkitRenderer, loadSystemFonts } from "../packages/pdfkit/src/node";
import { boxRadii, dilate, gaussianBlur } from "../packages/pdfkit/src/effect/filter";
import { composite } from "../packages/pdfkit/src/effect/raster";
import { breakOpportunities } from "../packages/pdfkit/src/lineBreak";
import { record, strokeOutline } from "../packages/pdfkit/src/path";
import { layoutText } from "../packages/pdfkit/src/text/layout";
import { createShaper } from "../packages/pdfkit/src/text/shaper";
import { populateWebpDeck } from "./visual/wasm/format-deck";

const PAGE = PAGE_SIZES.SCREEN_16X9.landscape;
/** The page height in points: PDFKit flips each text object about it, so a baseline `y` px from the top is written as `PAGE_HEIGHT_PT - y`. */
const PAGE_HEIGHT_PT = 405;
const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
const GIF = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7";
/** A 2x1 lossless WebP (ImageMagick 7.1), pdf-core's test picture. */
const WEBP = "data:image/webp;base64,UklGRhwAAABXRUJQVlA4TA8AAAAvAQAAAAcQ/Y/+ByKi/wEA";
const createDocument = () => new JSDOM("<!DOCTYPE html><html><head></head><body></body></html>").window.document;
const systemFonts = await loadSystemFonts();

/** The inflated content streams of every page, in page order (PDFKit starts each with its y flip). */
function pageContents(pdf: Uint8Array): string[] {
    const bytes = Buffer.from(pdf);
    const contents: string[] = [];

    for (const match of bytes.toString("latin1").matchAll(/stream\n/g)) {
        const start = (match.index ?? 0) + match[0].length;

        try {
            const text = inflateSync(bytes.subarray(start, bytes.indexOf("endstream", start))).toString("latin1");

            if (text.startsWith("1 0 0 -1 0 ")) {
                contents.push(text);
            }
        } catch {
            // a stream that is not deflated (or not a content stream): fonts and images
        }
    }

    return contents;
}

/** The dictionaries of the image XObjects in `pdf`, in object order. */
function imageObjects(pdf: Uint8Array): string[] {
    const images: string[] = [];

    for (const [image] of Buffer.from(pdf)
        .toString("latin1")
        .matchAll(/<<[^<>]*\/Subtype \/Image[^<>]*>>/g)) {
        if (!image.includes("/DeviceGray")) {
            images.push(image);
        }
    }

    return images;
}

/** Each drawn text object's position in CSS px from the slide's top left, and its font size in px. */
function texts(content: string): { x: number; baseline: number; size: number }[] {
    return [...content.matchAll(/1 0 0 1 ([\d.-]+) ([\d.-]+) Tm\n\/F\d+ ([\d.]+) Tf/g)].map(([, x, y, size]) => ({ x: Number(x), baseline: PAGE_HEIGHT_PT - Number(y), size: Number(size) }));
}

async function render(renderer: PdfkitRenderer, populate: (pres: PuppeteerGen) => void): Promise<Uint8Array> {
    const pres = new PuppeteerGen(PAGE, { renderer, createDocument });

    populate(pres);

    return pres.toPdf();
}

describe("PDFKit renderer", () => {
    test("places a line at the box's inset and the profile's baseline, at the run's size", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const pdf = await render(renderer, pres => pres.addSlide().addText("Hello", { x: 1, y: 1, w: 4, h: 1, fontFace: "Arial", fontSize: 18, margin: 0, valign: "top" }));

        // 1 in = 96 px; 18 pt = 24 px; the first baseline sits data-line-baseline (24 px for 1.2 x 24 px lines) below the top.
        expect(texts(pageContents(pdf)[0] ?? "")).toEqual([{ x: 96, baseline: 120, size: 24 }]);
        expect(renderer.warnings()).toEqual([]);
    });

    test("mirrors the text back inside a box flipped one way, about its text area's middle (#627)", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const options = { x: 1, y: 1, w: 4, h: 1, fontFace: "Arial", fontSize: 18, margin: 0, valign: "top" } as const;
        const pdf = await render(renderer, pres => {
            pres.addSlide().addText("Hello", { ...options, flipH: true });
            pres.addSlide().addText("Hello", { ...options, flipH: true, flipV: true });
        });
        const [mirrored = "", turned = ""] = pageContents(pdf);
        // The text area runs from 96 to 480 px, so x maps to 576 - x; under both flips the box only turns.
        const mirror = /\n-1 0 -?0 1 576 0 cm\n/;

        expect(mirrored).toMatch(mirror);
        expect(turned).not.toMatch(mirror);
        expect(texts(mirrored)).toEqual([{ x: 96, baseline: 120, size: 24 }]);
        expect(renderer.warnings()).toEqual([]);
    });

    test("wraps a paragraph onto line boxes one line-height apart, and centres each line", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const words = "one two three four five six seven eight nine ten eleven twelve";
        const pdf = await render(renderer, pres => pres.addSlide().addText(words, { x: 1, y: 1, w: 1.5, h: 2, fontFace: "Arial", fontSize: 12, margin: 0, valign: "top", align: "center" }));
        // Each word is its own text object; a line is the words on one baseline.
        const drawn = texts(pageContents(pdf)[0] ?? "");
        const baselines = [...new Set(drawn.map(text => text.baseline))];

        expect(baselines.length).toBeGreaterThan(2);
        // 12 pt = 16 px at 1.2 line spacing: 19.2 px per line.
        for (const [index, baseline] of baselines.entries()) {
            expect(baseline).toBeCloseTo((baselines[0] ?? 0) + index * 19.2, 3);
        }

        for (const baseline of baselines) {
            const xs: number[] = [];

            for (const text of drawn) {
                if (text.baseline === baseline) {
                    xs.push(text.x);
                }
            }

            expect(Math.min(...xs)).toBeGreaterThan(96);
        }
    });

    test("draws a stretched image over its box, clipped to it", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const pdf = await render(renderer, pres => pres.addSlide().addImage({ data: PNG, x: 2, y: 1, w: 1, h: 0.5 }));
        const content = pageContents(pdf)[0] ?? "";

        expect(content).toContain("192 96 96 48 re\nW n");
        // The image matrix maps the unit square onto the box, flipped to PDFKit's y-down space: bottom edge at 144 px.
        expect(content).toContain("96 0 0 -48 192 144 cm\n/I1 Do");
    });

    test("throws on what it cannot draw, or lists it in warnings() under onUnsupported: warn", async () => {
        const populate = (pres: PuppeteerGen) => pres.addSlide().addTable([[{ text: "cell", options: { textDirection: "vert" } }]], { x: 1, y: 1, w: 2 });

        await expect(render(await createPdfkitRenderer({ fonts: systemFonts }), populate)).rejects.toThrow(
            /The PDFKit renderer cannot draw everything in this presentation \(pass \{ onUnsupported: "warn" \} to render anyway\):\n- slide 1: object 'Table 0': vertical text is not drawn/,
        );
        const renderer = await createPdfkitRenderer({ fonts: systemFonts, onUnsupported: "warn" });
        const pdf = await render(renderer, populate);

        expect(Buffer.from(pdf).toString("latin1").startsWith("%PDF-")).toBe(true);
        expect(renderer.warnings()).toEqual(["slide 1: object 'Table 0': vertical text is not drawn"]);
    });

    describe("tables", () => {
        const solid = { type: "solid", pt: 3, color: "FF0000" } as const;
        const none = { type: "none" } as const;
        // Two 1 in columns, two 0.5 in rows, at (1 in, 1 in): grid lines at x 96/192/288 px and y 96/144/192 px.
        const grid = { x: 1, y: 1, colW: [1, 1], rowH: [0.5, 0.5], margin: 0, fontFace: "Arial", fontSize: 12 } satisfies PptxGenJS.TableProps;

        test("fills each cell's rect on the column grid, a merged cell over every column it spans", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            const pdf = await render(renderer, pres => pres.addSlide().addTable([[{ text: "", options: { colspan: 2, fill: { color: "00FF00" } } }], [{ text: "", options: { fill: { color: "0000FF" } } }, { text: "" }]], grid));
            const content = pageContents(pdf)[0] ?? "";

            expect(content).toContain("0 1 0 scn\n/Gs1 gs\n96 96 192 48 re\nf");
            expect(content).toContain("0 0 1 scn\n/Gs1 gs\n96 144 96 48 re\nf");
            expect(renderer.warnings()).toEqual([]);
        });

        test("strokes each shared grid edge once, a run of equal edges as one line centred on the grid line", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            const pdf = await render(renderer, pres =>
                pres.addSlide().addTable(
                    [
                        [
                            { text: "", options: { border: [solid, none, none, none] } },
                            { text: "", options: { border: [solid, none, none, none] } },
                        ],
                        [{ text: "" }, { text: "", options: { border: [none, none, none, { type: "dash", pt: 3, color: "0000FF" }] } }],
                    ],
                    { ...grid, border: none },
                ),
            );
            const strokes = [...(pageContents(pdf)[0] ?? "").matchAll(/([\d.]+) w\n(\d+) J\n(?:\[[^\]]*\] [\d.]+ d\n)?([\d.]+) ([\d.]+) m\n([\d.]+) ([\d.]+) l\nS/g)].map(([, width, cap, ...points]) => ({
                width: Number(width),
                cap: Number(cap),
                points: points.map(Number),
            }));

            // 3 pt = 4 px: the two top sides are one square-capped line across both columns; the dashed left side of the
            // second row's second cell is the shared vertical edge at x 192, butt-capped.
            expect(strokes).toEqual([
                { width: 4, cap: 2, points: [96, 96, 288, 96] },
                { width: 4, cap: 0, points: [192, 144, 192, 192] },
            ]);
        });

        test("places a cell's text at its margins, anchored in the row, and grows a row to its text", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            const pdf = await render(renderer, pres =>
                pres.addSlide().addTable(
                    [
                        [
                            { text: "Top", options: { valign: "top", margin: [6, 0, 0, 12] } },
                            { text: "Bottom", options: { valign: "bottom" } },
                        ],
                        [{ text: "one\ntwo\nthree\nfour", options: { colspan: 2 } }],
                    ],
                    grid,
                ),
            );
            const drawn = texts(pageContents(pdf)[0] ?? "");
            // 12 pt = 16 px Arial: 19.2 px line boxes with the baseline (data-line-baseline) 16 px below their top.
            const baseline = 16;

            // Margins are in pt: top 6 pt = 8 px, left 12 pt = 16 px.
            expect(drawn[0]).toEqual({ x: 112, baseline: 104 + baseline, size: 16 });
            // Bottom anchored in the 48 px first row: one 19.2 px line box ending at 144.
            expect(drawn[1]?.x).toBe(192);
            expect(drawn[1]?.baseline).toBeCloseTo(144 - 19.2 + baseline, 3);
            // Four 19.2 px lines are taller than the 48 px row minimum, so the second row grows to 76.8 px.
            expect(drawn.slice(2).map(text => text.baseline)).toEqual([0, 1, 2, 3].map(line => expect.closeTo(144 + line * 19.2 + baseline, 3)));
            expect(renderer.warnings()).toEqual([]);
        });
    });

    test(
        "draws shadows and glows within a per-render mask sample budget, as pdf-core does (GHSA-mc68-2mrj-jwhm)",
        async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts, onUnsupported: "warn" });
            // Shapes larger than the mask's longest side each take a full mask: the budget holds four of them.
            const shadowed = (count: number) => (pres: PuppeteerGen) => {
                const slide = pres.addSlide();

                for (let index = 0; index < count; index++) {
                    slide.addShape("rect", { x: 0, y: 0, w: 45, h: 45, fill: { color: "4472C4" }, shadow: { type: "outer", blur: 1, offset: 1, angle: 45, color: "000000", opacity: 0.5 } });
                }
            };

            await render(renderer, shadowed(5));
            expect(renderer.warnings()).toEqual(["slide 1: object 'Shape 4': shadows and glows past 67108864 mask samples per render not rendered"]);
            // The budget is per render.
            await render(renderer, shadowed(1));
            expect(renderer.warnings()).toEqual([]);
        },
        { timeout: 60_000 },
    );

    test("cuts an effect outline's dashes only where a mask can show them (GHSA-mc68-2mrj-jwhm)", () => {
        const style = { width: 1, cap: "butt", join: "miter", miterLimit: 10 } as const;
        const line = (length: number) => [
            {
                points: [
                    [0, 0],
                    [length, 0],
                ] as [number, number][],
                closed: false,
            },
        ];

        // A 2-unit pattern over 100 units: 50 dashes, one quadrilateral each.
        expect(strokeOutline(line(100), { ...style, dashes: [1, 1] })).toHaveLength(50);
        // A pattern finer than a mask sample is stroked solid: one quadrilateral, not 500,000.
        expect(strokeOutline(line(100), { ...style, dashes: [0.0001, 0.0001] })).toHaveLength(1);
        // Past tiny-skia's 1,000,000 dashes the stroke casts nothing, as in pdf-core.
        expect(strokeOutline(line(4_000_000), { ...style, dashes: [1, 1] })).toEqual([]);
    });

    test("draws a hairline dotted outline's shadow in bounded time", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const start = performance.now();

        await render(renderer, pres =>
            pres.addSlide().addShape("rect", { x: 0.5, y: 0.5, w: 9, h: 4.5, line: { color: "000000", width: 0.001, dashType: "sysDot" }, shadow: { type: "outer", blur: 4, offset: 4, angle: 45, color: "000000", opacity: 0.5 } }),
        );
        // About 3 s before the fix, about 25 ms after.
        expect(performance.now() - start).toBeLessThan(1000);
    });

    describe("pictures", () => {
        test("draws a GIF's and a WebP's first frame, embedded as PNG", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            const pdf = await render(renderer, pres => {
                pres.addSlide().addImage({ data: GIF, x: 1, y: 1, w: 1, h: 1 });
                pres.addSlide().addImage({ data: WEBP, x: 1, y: 1, w: 2, h: 1 });
            });

            expect(
                imageObjects(pdf)
                    .map(image => [/\/Width (\d+)/.exec(image)?.[1], /\/Height (\d+)/.exec(image)?.[1]])
                    .sort(),
            ).toEqual([
                ["1", "1"],
                ["2", "1"],
            ]);
            expect(pageContents(pdf).every(content => /\/I\d+ Do/.test(content))).toBe(true);
            expect(renderer.warnings()).toEqual([]);
        });

        test("writes the same bytes for the same deck: pictures are embedded in drawing order (#424)", async () => {
            // The WebP deck's pictures go to PDFKit as PNGs with alpha, whose embedding PDFKit finishes after an
            // asynchronous inflate: unordered, the object numbers followed whichever inflate ended first.
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            // The Node entry reads the deck's picture files.
            const html = await populateWebpDeck(new NodePuppeteerGen()).toHtml();
            const first = await renderer.render(html);
            const second = await renderer.render(html);

            expect(imageObjects(first).length).toBeGreaterThan(2);
            expect(Buffer.from(second).equals(Buffer.from(first))).toBe(true);
        });

        test("refuses a GIF whose first frame does not fill its screen, as pdf-core does", async () => {
            // GIF89a: a 2x1 screen, a two-colour table, then a 1x1 first frame at (1, 0).
            const gif = Buffer.concat([Buffer.from("GIF89a\x02\x00\x01\x00\x80\x00\x00\xff\x00\x00\x00\x00\xff", "latin1"), Buffer.from([0x2c, 1, 0, 0, 0, 1, 0, 1, 0, 0, 2, 2, 0x44, 0x0a, 0, 0x3b])]);
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });

            await expect(render(renderer, pres => pres.addSlide().addImage({ data: `data:image/gif;base64,${gif.toString("base64")}`, x: 1, y: 1, w: 1, h: 1 }))).rejects.toThrow(
                "the GIF's first frame (1x1 at 1,0) does not fill its 2x1 screen; save it with a full first frame",
            );
        });

        test("draws an SVG picture as vector paths and reports what is outside the SVG subset", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts, onUnsupported: "warn" });
            const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 20 10"><title>t</title><rect width="10" height="10" fill="red"/><circle cx="15" cy="5" r="4" fill="none" stroke="#0000FF" stroke-width="2"/><text>no</text></svg>`;
            const pdf = await render(renderer, pres => pres.addSlide().addImage({ data: `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`, x: 1, y: 1, w: 2, h: 1 }));

            expect(imageObjects(pdf)).toEqual([]);
            const [content = ""] = pageContents(pdf);

            // The viewBox maps 20 x 10 units onto the 192 x 96 px picture: 9.6 px per unit.
            expect(content).toContain("9.6 0 0 9.6 96 96 cm");
            expect(content).toContain("1 0 0 scn\n/Gs1 gs\n0 0 10 10 re\nf");
            expect(content).toContain("0 0 1 SCN");
            expect(renderer.warnings()).toEqual(["slide 1: object 'Image 0': <text> in an SVG image is not drawn"]);
        });

        test("fades a translucent SVG picture as one group, as Chromium composites an img", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><rect width="10" height="10" fill="red"/><circle cx="5" cy="5" r="3" fill="white"/></svg>`;
            const pdf = await render(renderer, pres => pres.addSlide().addImage({ data: `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`, x: 1, y: 1, w: 1, h: 1, transparency: 50 }));
            const [content = ""] = pageContents(pdf);

            // Both shapes are in the group, painted once with the picture's alpha: the circle hides the rect under it.
            expect(content).toContain("/FG1 Do");
            expect(content).not.toContain("0 0 10 10 re");
            expect(Buffer.from(pdf).toString("latin1")).toMatch(/\/Subtype \/Form\n\/BBox \[[^\]]*\]\n\/Group <<\n\/Type \/Group\n\/S \/Transparency\n>>/);
            expect(renderer.warnings()).toEqual([]);
        });

        test("names the invalid opacity an SVG picture's style declares, not its attribute (#430)", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts, onUnsupported: "warn" });
            const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10"><rect width="10" height="10" fill="red" fill-opacity="0.5" style="fill-opacity: half"/><rect width="5" height="5" style="stroke-opacity: none"/></svg>`;

            await render(renderer, pres => pres.addSlide().addImage({ data: `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`, x: 1, y: 1, w: 1, h: 1 }));
            expect(renderer.warnings()).toEqual(["slide 1: object 'Image 0': fill-opacity 'half' in an SVG image is not a number; inherited instead", "slide 1: object 'Image 0': stroke-opacity 'none' in an SVG image is not a number; inherited instead"]);
        });

        test("fails the render, not the process, on a corrupt PNG whose pixels PDFKit would decode (GHSA-mc68-2mrj-jwhm)", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            const chunk = (type: string, data: Buffer) => {
                const head = Buffer.alloc(8);
                const crc = Buffer.alloc(4);

                head.writeUInt32BE(data.length, 0);
                head.write(type, 4, "latin1");
                crc.writeUInt32BE(crc32(Buffer.concat([Buffer.from(type, "latin1"), data])), 0);

                return Buffer.concat([head, data, crc]);
            };
            // A 4x4 RGBA PNG: IDAT is a valid zlib stream of filter bytes 9 (no such filter), or 68 good bytes.
            const png = (idat: Buffer) => {
                const ihdr = Buffer.from([0, 0, 0, 4, 0, 0, 0, 4, 8, 6, 0, 0, 0]);

                return `data:image/png;base64,${Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), chunk("IHDR", ihdr), chunk("IDAT", idat), chunk("IEND", Buffer.alloc(0))]).toString("base64")}`;
            };

            await expect(render(renderer, pres => pres.addSlide().addImage({ data: png(deflateSync(Buffer.alloc(68, 9))), x: 1, y: 1, w: 1, h: 1 }))).rejects.toThrow("cannot decode image/png data");
            await expect(render(renderer, pres => pres.addSlide().addImage({ data: png(Buffer.from([0x78, 0x9c, 0xff, 0xff, 0xff, 0x00])), x: 1, y: 1, w: 1, h: 1 }))).rejects.toThrow("cannot decode image/png data");
            expect((await render(renderer, pres => pres.addSlide().addImage({ data: png(deflateSync(Buffer.alloc(68))), x: 1, y: 1, w: 1, h: 1 }))).length).toBeGreaterThan(0);
        });

        test("reports image formats it does not decode", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts, onUnsupported: "warn" });

            await render(renderer, pres => pres.addSlide().addImage({ data: "data:image/bmp;base64,Qk0=", x: 1, y: 1, w: 1, h: 1 }));
            expect(renderer.warnings()).toEqual(["slide 1: object 'Image 0': image format 'image/bmp' is not drawn (not a PNG, JPEG, GIF, WebP or SVG image)"]);
        });
    });

    test("draws a strike in the text colour when the underline has its own (#430)", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const pdf = await render(renderer, pres => pres.addSlide().addText("Hi", { x: 1, y: 1, w: 3, h: 1, fontFace: "Arial", color: "0000FF", underline: { style: "sng", color: "FF0000" }, strike: "sngStrike" }));

        // The underline's band in its red, then the strike's in the run's blue, as PowerPoint and pdf-core draw them.
        expect(pageContents(pdf)[0]).toMatch(/ re\n\/DeviceRGB cs\n1 0 0 scn\n\/Gs1 gs\nf\n[\d. ]+ re\n\/DeviceRGB cs\n0 0 1 scn\n\/Gs1 gs\nf\nQ\n$/);
    });

    test("reports a -webkit-text-stroke with a width or a colour alone, the profile's form being both (#430)", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts, onUnsupported: "warn" });
        const pres = new PuppeteerGen(PAGE, { renderer, createDocument });

        pres.addSlide().addText("Hi", { x: 1, y: 1, w: 3, h: 1, fontFace: "Arial", outline: { size: 1, color: "00FF00" } });
        const document = await pres.toHtml();

        for (const value of ["2px", "#00FF00"]) {
            const pdf = await renderer.render({ ...document, html: document.html.replace(/-webkit-text-stroke: [^;"]*/, `-webkit-text-stroke: ${value}`) });

            expect(pageContents(pdf)[0]).not.toContain("1 Tr");
            expect(renderer.warnings()).toEqual([`slide 1: object 'Text 0': -webkit-text-stroke '${value}' is not drawn`]);
        }
    });

    test("drops a tab no face draws, without counting it as text in a missing face (#430)", () => {
        const color = { rgb: [0, 0, 0], alpha: 1 } as const;
        const run = { text: "a\tb", face: { family: "Nothing", fallbacks: [], bold: false, italic: false }, size: 16, color, letterSpacing: 0, kerning: true, shift: 0, softBreakBefore: false, synthesis: { boldStroke: 0, italicShear: 0 } };
        const paragraph = { runs: [run], align: "left", rtl: false, lineHeight: 19.2, baseline: 15, markBaseline: 15, breakBaseline: 15, spaceBefore: 0, spaceAfter: 0, tabSize: 48, tabStops: [], indent: 0, textIndent: 0 } as const;
        const missing: string[] = [];
        // No faces at all: the renderer refuses an empty font list, so only a direct call reaches an empty chain.
        const layout = layoutText({ paragraphs: [paragraph], padding: [0, 0, 0, 0], anchor: "flex-start", wrap: true, fit: { kind: "none" }, mirrored: false }, 200, 1, createShaper([]), missed => missing.push(missed.text));

        expect(layout.lines.flatMap(line => line.fragments)).toEqual([]);
        // "a" and "b" are text in a missing face; the tab between them is not.
        expect(missing).toEqual(["a\tb", "a\tb"]);
    });

    test("records a quadratic as the cubic with its curve, not PDFKit's `v`", () => {
        const doc = new PDFDocument({ autoFirstPage: false });

        // The controls sit two thirds of the way from each end to the quadratic's control (30, 60).
        expect(record(doc, pen => pen.path("M 0 0 Q 30 60 90 0"))).toEqual([
            ["M", 0, 0],
            ["C", 20, 40, 50, 40, 90, 0],
        ]);
    });

    describe("line heads and effects", () => {
        test("fills each line head in the line's colour at its end, turned along the line", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            const pdf = await render(renderer, pres => pres.addSlide().addShape("line", { x: 1, y: 1, w: 2, h: 0, line: { color: "FF0000", width: 2, beginArrowType: "oval", endArrowType: "triangle" } }));
            const [content = ""] = pageContents(pdf);

            // The page background, the stroke, then two filled heads; the line runs left to right, so the end head is not
            // turned (cos 1, sin 0) and the start head points back (cos -1).
            expect(content.match(/^f$/gm)?.length).toBe(3);
            expect(content).toMatch(/\n-1 0 -?0 -1 0 0 cm\n/);
            expect(content).toMatch(/\n1 0 -?0 1 192 0 cm\n/);
            expect(renderer.warnings()).toEqual([]);
        });

        test("puts an outer shadow under its shape, as a blurred image whose soft mask is the silhouette", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            const pdf = await render(renderer, pres => pres.addSlide().addShape("rect", { x: 1, y: 1, w: 1, h: 1, fill: { color: "FF0000" }, shadow: { type: "outer", blur: 6, offset: 4, angle: 45, color: "000000", opacity: 0.5 } }));
            const [image = ""] = imageObjects(pdf).filter(object => object.includes("/SMask"));

            expect(image).toContain("/Interpolate true");
            expect(image).toContain("/ColorSpace /DeviceRGB");
            const [content = ""] = pageContents(pdf);

            expect(content.indexOf("/FX1 Do")).toBeGreaterThan(-1);
            expect(content.indexOf("/FX1 Do")).toBeLessThan(content.indexOf("1 0 0 scn"));
            expect(renderer.warnings()).toEqual([]);
        });

        test("glows text under its glyphs", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts });
            const pdf = await render(renderer, pres => pres.addSlide().addText("Glow", { x: 1, y: 1, w: 3, h: 1, glow: { size: 8, opacity: 0.5, color: "FFC000" } }));
            const [content = ""] = pageContents(pdf);

            expect(content.indexOf("/FX1 Do")).toBeGreaterThan(-1);
            expect(content.indexOf("/FX1 Do")).toBeLessThan(content.indexOf("BT"));
            expect(renderer.warnings()).toEqual([]);
        });

        test("draws a picture's inner shadow over it, and reports one on a shape as pdf-core does", async () => {
            const renderer = await createPdfkitRenderer({ fonts: systemFonts, onUnsupported: "warn" });
            const picture = new PuppeteerGen(PAGE, { renderer, createDocument, target: "powerpoint" });

            picture.addSlide().addImage({ data: PNG, x: 1, y: 1, w: 1, h: 1, shadow: { type: "inner", blur: 6, offset: 4, angle: 45, color: "000000", opacity: 0.5 } });
            const [content = ""] = pageContents(await picture.toPdf());

            expect(content.indexOf("/FX1 Do")).toBeGreaterThan(content.indexOf("/I1 Do"));
            const shape = new PuppeteerGen(PAGE, { renderer, createDocument });

            shape.addSlide().addShape("rect", { x: 1, y: 1, w: 1, h: 1, fill: { color: "FF0000" }, shadow: { type: "outer", blur: 6, offset: 4, angle: 45, color: "000000", opacity: 0.5 } });
            const { html, pageSize } = await shape.toHtml();
            const pdf = await renderer.render({ html: html.replace('<svg class="shape-svg"', '<svg class="shape-svg" data-shadow-inset=""'), pageSize });

            expect(pageContents(pdf).join("")).not.toContain(" Do");
            expect(renderer.warnings()).toEqual(["slide 1: object 'Shape 0': inner shadow not rendered"]);
        });

        test("blurs and dilates a mask as pdf-core's filters do", () => {
            const [w, h] = [21, 21];
            const mask = new Float32Array(w * h);

            composite(
                [
                    [8, 8, 13, 8],
                    [13, 8, 13, 13],
                    [13, 13, 8, 13],
                    [8, 13, 8, 8],
                ],
                1,
                w,
                h,
                new Float32Array(w * h + 2),
                mask,
            );
            const sum = (values: Float32Array) => {
                let total = 0;

                for (const value of values) {
                    total += value;
                }

                return total;
            };

            // The 5 x 5 square is covered exactly.
            expect(sum(mask)).toBeCloseTo(25, 4);
            expect(mask[10 * w + 10]).toBe(1);
            gaussianBlur(mask, w, h, 1.5);
            // A blur away from the edges moves coverage, never adds or loses it.
            expect(sum(mask)).toBeCloseTo(25, 2);
            expect(mask[10 * w + 10]).toBeLessThan(1);
            const point = new Float32Array(w * h);

            point[10 * w + 10] = 1;
            dilate(point, w, h, 3);
            // Grown by 3 px with an anti-aliased edge, half a px each side of it.
            expect(point[10 * w + 12]).toBe(1);
            expect(point[10 * w + 13]).toBe(0.5);
            expect(point[10 * w + 14]).toBe(0);
            expect(point[10 * w + 10]).toBe(1);
            // The three box blurs' cascade has about the Gaussian's standard deviation (pdf-core's own test).
            for (const sigma of [1, 2.5, 3.7, 7, 20]) {
                let variance = 0;

                for (const radius of boxRadii(sigma)) {
                    variance += ((2 * radius + 1) ** 2 - 1) / 12;
                }

                expect(Math.abs(Math.sqrt(variance) - sigma)).toBeLessThan(0.25 * Math.max(sigma, 1));
            }
        });
    });

    test("lists a face no font provides, with the face drawn instead", async () => {
        const missing: unknown[] = [];
        const renderer = await createPdfkitRenderer({ fonts: systemFonts, onMissingFont: list => missing.push(...list) });

        await render(renderer, pres => pres.addSlide().addText("x", { x: 1, y: 1, w: 1, h: 1, fontFace: "No Such Face" }));
        expect(missing).toEqual([{ face: "No Such Face", objects: ["slide 1 'Text 0'"], usedFace: expect.any(String) }]);
        expect(renderer.missingFonts()).toEqual(missing as never);
    });

    test("writes the same bytes for the same deck, produced by PdfLapas", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const populate = (pres: PuppeteerGen) => pres.addSlide().addText("Same", { x: 1, y: 1, w: 2, h: 1 });
        const pdf = Buffer.from(await render(renderer, populate));

        expect(pdf.equals(Buffer.from(await render(renderer, populate)))).toBe(true);
        // PDFKit derives the file ID from a date: the fixed epoch keeps it stable, and the Info dictionary has no date.
        const source = pdf.toString("latin1");

        expect(source).toMatch(/\/Producer \d+ 0 R\n\/Creator \d+ 0 R\n>>/);
        expect(source).toContain("(PdfLapas)");
        expect(source).not.toContain("/CreationDate");
    });

    test("writes the deck's PDF properties, a date alone as both dates, and refuses a date that is not UTC ISO", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const pres = new PuppeteerGen(PAGE, { renderer, createDocument, pdfMetadata: { keywords: "a, b", creator: "Maker", producer: "Press", creationDate: new Date("2026-11-19T12:34:56.789Z") } });

        pres.addSlide();
        const source = Buffer.from(await pres.toPdf()).toString("latin1");

        for (const value of ["(a, b)", "(Maker)", "(Press)"]) {
            expect(source).toContain(value);
        }

        expect(source.match(/\(D:20261119123456Z\)/g)?.length).toBe(2);
        const document = await pres.toHtml();
        const html = document.html.replace("2026-11-19T12:34:56.789Z", "2026-11-19T12:34:56+01:00");

        await expect(renderer.render({ ...document, html })).rejects.toThrow('"2026-11-19T12:34:56+01:00" is not a UTC date');
    });

    test("reads the PowerPoint target's word spans, fallback faces and non-breaking hyphen", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const pres = new PuppeteerGen(PAGE, { renderer, createDocument, target: "powerpoint" });

        pres.addSlide().addText("Lorem ipsum dolor sit amet non\u2011breaking", { x: 1, y: 1, w: 6, h: 1, fontFace: "Arial", fontSize: 10, margin: 0 });
        const document = await pres.toHtml();

        expect(document.html).toContain("font-family: &quot;Arial&quot;, &quot;Cambria Math&quot;");
        expect(document.html).toContain('data-non-breaking-hyphen="hyphen-minus"');
        // Word spans as the profile writes them under 12pt for a face whose advances are known.
        const html = document.html.replace(">Lorem ipsum", '><span style="letter-spacing: 1.5px;">Lorem</span> <span style="letter-spacing: 0.5px;">ipsum</span>');
        const content = pageContents(await renderer.render({ ...document, html }))[0] ?? "";

        expect(renderer.warnings()).toEqual([]);
        // Each word at its own spacing (Tc); PDFKit writes none for the bare text's, the run's 0.
        expect([...content.matchAll(/([\d.-]+) Tc/g)].map(match => match[1])).toEqual(["1.5", "0.5"]);
        const doc = new PuppeteerGen(PAGE, { renderer, createDocument, target: "powerpoint" });

        doc.addSlide().addText("x", { x: 1, y: 1, w: 1, h: 1 });
        const bare = await doc.toHtml();

        await expect(renderer.render({ ...bare, html: bare.html.replace("font-family: &quot;", "font-family: Arial, &quot;") })).rejects.toThrow("is not one or more quoted families");
    });

    test("reads the LibreOffice target's non-breaking hyphen, U+2010, and rejects an unknown one", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts });
        const pres = new PuppeteerGen(PAGE, { renderer, createDocument });

        pres.addSlide().addText("non‑breaking", { x: 1, y: 1, w: 6, h: 1, fontFace: "Arial", fontSize: 10, margin: 0 });
        const document = await pres.toHtml();

        expect(document.html).toContain('data-non-breaking-hyphen="hyphen"');
        await renderer.render(document);
        expect(renderer.warnings()).toEqual([]);
        await expect(renderer.render({ ...document, html: document.html.replace('data-non-breaking-hyphen="hyphen"', 'data-non-breaking-hyphen="dash"') })).rejects.toThrow("attribute 'data-non-breaking-hyphen' 'dash' is not hyphen-minus or hyphen");
    });

    test("kerns a pair across a break opportunity, as one shaped run", async () => {
        const fonts = await loadSystemFonts({ match: /^LiberationSans-Regular\.ttf$/ });
        const [file] = fonts;

        if (!file) {
            throw new Error("Liberation Sans is not installed");
        }

        const font = createFont(Buffer.from(file)) as Font; // `as`: a .ttf holds one font, not a collection
        const renderer = await createPdfkitRenderer({ fonts });
        const pdf = await render(renderer, pres => pres.addSlide().addText("Same API", { x: 1, y: 1, w: 4, h: 1, fontFace: "Arial", fontSize: 18, margin: 0 }));
        const [same, api] = texts(pageContents(pdf)[0] ?? "");
        // " A" is a kern pair: "API" starts where shaping the whole run puts it, not after an unkerned "Same ".
        const kerned = (font.layout("Same API").advanceWidth - font.layout("API").advanceWidth) / font.unitsPerEm;

        expect(font.layout("Same ").advanceWidth / font.unitsPerEm).not.toBeCloseTo(kerned, 4);
        expect((api?.x ?? 0) - (same?.x ?? 0)).toBeCloseTo(kerned * 24, 3);
    });

    describe("underline and strike geometry (docs/HTML_PROFILE.md)", () => {
        // "Ab cd" at 24px underlined on the LibreOffice target, its HTML rewritten by `edit`; its band and the drawn rects.
        const underlined = async (edit: (html: string) => string = html => html, options: PptxGenJS.TextPropsOptions = {}) => {
            const renderer = await createPdfkitRenderer({ fonts: await loadSystemFonts({ match: /^LiberationSans-Regular\.ttf$/ }), onUnsupported: "warn" });
            const pres = new PuppeteerGen(PAGE, { renderer, createDocument });

            pres.addSlide().addText("Ab cd", { x: 1, y: 1, w: 4, h: 1, fontFace: "Arial", fontSize: 18, underline: { style: "sng" }, margin: 0, valign: "top", ...options });
            const document = await pres.toHtml();
            const offset = Number(/text-underline-offset: ([\d.-]+)em/.exec(document.html)?.[1]);
            const thickness = Number(/text-decoration-thickness: ([\d.]+)em/.exec(document.html)?.[1]);
            const content = pageContents(await renderer.render({ ...document, html: edit(document.html) }))[0] ?? "";
            const [text] = texts(content);
            // The first rect is the slide background.
            const rects = [...content.matchAll(/([\d.-]+) ([\d.-]+) ([\d.-]+) ([\d.-]+) re/g)].slice(1).map(match => match.slice(1).map(Number));

            return { content, rects, x: text?.x ?? 0, top: (text?.baseline ?? 0) + offset * 24, t: thickness * 24, warnings: renderer.warnings() };
        };

        test("draws the band the HTML states with the target's pieces: LibreOffice's double lines", async () => {
            const double = await underlined(html => html, { underline: { style: "dbl" } });

            expect(double.rects.map(rect => rect[1])).toEqual([expect.closeTo(double.top - 0.46 * double.t, 3), expect.closeTo(double.top + 1.46 * double.t, 3)]);
            expect(double.rects.map(rect => rect[3])).toEqual([expect.closeTo(0.64 * double.t, 3), expect.closeTo(0.64 * double.t, 3)]);
            expect(double.warnings).toEqual([]);
        });

        test("cuts a line into dashes from the run's start, the last one cut at its end", async () => {
            const solid = await underlined();
            const [, , width = 0] = solid.rects[0] ?? [];
            const dashed = await underlined(html => html.replace('<span class="text-run"', '<span class="text-run" data-underline="line(0 1 4 2)"'));
            const t = dashed.t;

            expect(dashed.rects.length).toBe(Math.ceil(width / (6 * t)));
            expect(dashed.rects[1]?.[0]).toBeCloseTo((dashed.rects[0]?.[0] ?? 0) + 6 * t, 3);
            const last = dashed.rects.at(-1) ?? [];

            expect((last[0] ?? 0) + (last[2] ?? 0)).toBeCloseTo((solid.rects[0]?.[0] ?? 0) + width, 3);
            expect(last[2]).toBeLessThanOrEqual(4 * t + 1e-6);
        });

        test("strokes a wave and fills a zigzag around the band", async () => {
            const wave = await underlined(html => html.replace('<span class="text-run"', '<span class="text-run" data-underline="wave(0.5 2 4 0.03px)"'));

            // Cubic half-waves stroked 0.03px wide: no filled band.
            expect(wave.rects).toEqual([]);
            expect(wave.content).toMatch(/0\.03 w\n/);
            expect(wave.content).toMatch(/ c\n[^]*\nS\n/);
            const zigzag = await underlined(html => html.replace('<span class="text-run"', '<span class="text-run" data-underline="zigzag(-0.5 2 0.16px)"'));

            expect(zigzag.rects).toEqual([]);
            expect(zigzag.content).toMatch(/ l\nh\n[^]*\nf\n/);
            // A unitless line is in thicknesses: the same zigzag as its px length.
            const inBand = await underlined(html => html.replace('<span class="text-run"', `<span class="text-run" data-underline="zigzag(-0.5 2 ${0.16 / zigzag.t})"`));

            expect(inBand.content).toBe(zigzag.content);
            expect(inBand.warnings).toEqual([]);
        });

        test("reports a malformed piece or strike band and draws one stripe filling the face's band", async () => {
            const malformed = await underlined(html => html.replace('<span class="text-run"', '<span class="text-run" data-underline="line(0)" data-strike-band="1em"'), { strike: "sngStrike" });

            expect(malformed.warnings).toEqual([
                "slide 1: object 'Text 0': attribute 'data-strike-band' '1em' is not drawn: it is '<top>em <thickness>em' with a positive thickness (drawn with the face's strikeout band)",
                "slide 1: object 'Text 0': attribute 'data-underline' 'line(0)' is not drawn: 'line(0)' is not line(<top> <height>) with a height and an even count of dash lengths of at least 0.01 (drawn as line(0 1))",
            ]);
            expect(malformed.rects.map(rect => rect[1])[0]).toBeCloseTo(malformed.top, 3);
        });
    });

    test("draws the PowerPoint target's double and heavy underlines and one averaged band per line, as PowerPoint does (#379)", async () => {
        const fonts = await loadSystemFonts({ match: /^LiberationSans-Regular\.ttf$/ });
        const [file] = fonts;

        if (!file) {
            throw new Error("Liberation Sans is not installed");
        }

        const font = createFont(Buffer.from(file)) as Font; // `as`: a .ttf holds one font, not a collection
        const renderer = await createPdfkitRenderer({ fonts });
        const underlined = async (runs: PptxGenJS.TextProps[]) => {
            const pres = new PuppeteerGen(PAGE, { renderer, createDocument, target: "powerpoint" });

            pres.addSlide().addText(runs, { x: 1, y: 1, w: 8, h: 1, fontFace: "Arial", margin: 0, valign: "top" });
            const content = pageContents(await pres.toPdf())[0] ?? "";
            // The first rect is the slide background.
            const rects = [...content.matchAll(/([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) re/g)].slice(1).map(match => match.slice(1).map(Number));

            return { baseline: texts(content)[0]?.baseline ?? 0, lefts: rects.map(rect => rect[0] ?? 0), tops: rects.map(rect => rect[1] ?? 0), heights: rects.map(rect => rect[3] ?? 0) };
        };
        // The post band at `px`: its top below the baseline and its thickness.
        const band = (px: number) => [(-font.underlinePosition * px) / font.unitsPerEm, (font.underlineThickness * px) / font.unitsPerEm] as const;
        const [p, t] = band(24);
        const double = await underlined([{ text: "Double", options: { fontSize: 18, underline: { style: "dbl" } } }]);

        expect(double.tops).toEqual([expect.closeTo(double.baseline + p - t, 2), expect.closeTo(double.baseline + p + t / 2, 2)]);
        expect(double.heights).toEqual([expect.closeTo(t / 2, 2), expect.closeTo(t / 2, 2)]);
        const heavy = await underlined([{ text: "Heavy", options: { fontSize: 18, underline: { style: "heavy" } } }]);

        expect(heavy.tops).toEqual([expect.closeTo(heavy.baseline + p - t, 2)]);
        expect(heavy.heights).toEqual([expect.closeTo(1.5 * t, 2)]);
        // A 12pt then a 24pt run: one band for the line, between the two runs' own, the same under both.
        const mixed = await underlined([
            { text: "Small run ", options: { fontSize: 12, underline: { style: "sng" } } },
            { text: "Large run", options: { fontSize: 24, underline: { style: "sng" } } },
        ]);

        expect(new Set(mixed.tops).size).toBe(1);
        expect(new Set(mixed.heights).size).toBe(1);
        expect(mixed.heights[0]).toBeGreaterThan(band(16)[1]);
        expect(mixed.heights[0]).toBeLessThan(band(32)[1]);
        // Dashes under a 24pt then a 12pt run run on as one pattern of 3T on, 2T off in the line's band; a colour change
        // starts a new one at its run (#397).
        const dashed = await underlined([
            { text: "Large ", options: { fontSize: 24, underline: { style: "dash" } } },
            { text: "small ", options: { fontSize: 12, underline: { style: "dash" } } },
            { text: "other", options: { fontSize: 12, underline: { style: "dash", color: "C00000" } } },
        ]);
        const period = 5 * (dashed.heights[0] ?? 0);
        const restarts = [];

        for (const [index, left] of dashed.lefts.entries()) {
            const previous = dashed.lefts[index - 1];

            if (previous !== undefined && Math.abs(left - previous - period) > 0.01) {
                restarts.push(index);
            }
        }

        expect(restarts.length).toBe(1);
        expect(renderer.warnings()).toEqual([]);
    });

    test("draws a fontMetrics entry's underline band in place of the face's post band on the PowerPoint target (#391)", async () => {
        const fonts = await loadSystemFonts({ match: /^LiberationSans-Regular\.ttf$/ });
        const [file] = fonts;

        if (!file) {
            throw new Error("Liberation Sans is not installed");
        }

        const font = createFont(Buffer.from(file)) as Font; // `as`: a .ttf holds one font, not a collection
        const renderer = await createPdfkitRenderer({ fonts });
        const arial = { unitsPerEm: font.unitsPerEm, winAscent: 1854, winDescent: 434 };
        const underlined = async (underline?: { position: number; thickness: number }, style: PptxGenJS.TextPropsOptions["underline"] = { style: "sng" }) => {
            const fontMetrics = { Arial: { ...arial, ...(underline ? { underline } : {}) } };
            const pres = new PuppeteerGen(PAGE, { renderer, createDocument, target: "powerpoint", fontMetrics });

            pres.addSlide().addText("Band", { x: 1, y: 1, w: 8, h: 1, fontFace: "Arial", fontSize: 18, underline: style, margin: 0, valign: "top" });
            const content = pageContents(await pres.toPdf())[0] ?? "";
            // The first rect is the slide background.
            const rects = [...content.matchAll(/([\d.]+) ([\d.]+) ([\d.]+) ([\d.]+) re/g)].slice(1).map(match => match.slice(1).map(Number));

            return { content, baseline: texts(content)[0]?.baseline ?? 0, tops: rects.map(rect => rect[1] ?? 0), heights: rects.map(rect => rect[3] ?? 0) };
        };
        // 45 and 72 of 2048 units at 24px.
        const [p, t] = [(45 * 24) / font.unitsPerEm, (72 * 24) / font.unitsPerEm];
        const band = await underlined({ position: -45, thickness: 72 });

        expect(band.tops).toEqual([expect.closeTo(band.baseline + p, 2)]);
        expect(band.heights).toEqual([expect.closeTo(t, 2)]);
        const heavy = await underlined({ position: -45, thickness: 72 }, { style: "heavy" });

        expect(heavy.tops).toEqual([expect.closeTo(heavy.baseline + p - t, 2)]);
        expect(heavy.heights).toEqual([expect.closeTo(1.5 * t, 2)]);
        // The face's own post values, given explicitly, draw what the post band draws (to the 6 decimals of em the HTML writes).
        const post = await underlined({ position: font.underlinePosition, thickness: font.underlineThickness });
        const face = await underlined();

        expect(post.tops).toEqual(face.tops.map(top => expect.closeTo(top, 4)));
        expect(post.heights).toEqual(face.heights.map(height => expect.closeTo(height, 4)));
        expect(renderer.warnings()).toEqual([]);
    });

    test("puts text on explicit tab stops: a right stop ends it there, a decimal stop puts its '.' there", async () => {
        const fonts = await loadSystemFonts({ match: /^LiberationSans-Regular\.ttf$/ });
        const [file] = fonts;

        if (!file) {
            throw new Error("Liberation Sans is not installed");
        }

        const font = createFont(Buffer.from(file)) as Font; // `as`: a .ttf holds one font, not a collection
        const width = (text: string) => (font.layout(text).advanceWidth / font.unitsPerEm) * 24;
        const renderer = await createPdfkitRenderer({ fonts });
        const tabStops = [
            { position: 2, alignment: "r" },
            { position: 3, alignment: "dec" },
        ] as PptxGenJS.TextPropsOptions["tabStops"]; // `as`: PptxGenJS types the alignment as a string union the literals widen from
        const pdf = await render(renderer, pres => pres.addSlide().addText("a\tbb\t12.5", { x: 1, y: 1, w: 5, h: 1, fontFace: "Arial", fontSize: 18, margin: 0, tabStops }));

        expect(renderer.warnings()).toEqual([]);
        const [, right, decimal] = texts(pageContents(pdf)[0] ?? "");

        // Stops are measured from the text area's left edge: the box's x, 96 px.
        expect((right?.x ?? 0) + width("bb")).toBeCloseTo(96 + 192, 2);
        expect((decimal?.x ?? 0) + width("12")).toBeCloseTo(96 + 288, 2);
    });

    test("reports text on a right tab stop that wraps, which it aligns by its whole width (#460)", async () => {
        const renderer = await createPdfkitRenderer({ fonts: await loadSystemFonts({ match: /^LiberationSans-Regular\.ttf$/ }), onUnsupported: "warn" });
        const tabStops = [{ position: 1.5, alignment: "r" }] as PptxGenJS.TextPropsOptions["tabStops"]; // `as`: as above
        const options = { x: 1, y: 1, w: 2, h: 1, fontFace: "Arial", fontSize: 18, margin: 0, tabStops };

        await render(renderer, pres => pres.addSlide().addText("a\tbb", options));
        expect(renderer.warnings()).toEqual([]);
        await render(renderer, pres => pres.addSlide().addText("a\tbb cc dd ee ff gg hh ii jj kk", options));
        expect(renderer.warnings()).toEqual(["slide 1: object 'Text 0': text on a right, center or decimal tab stop wraps; it is aligned by its whole width, not by the part on the stop's line"]);
    });

    test("refuses elements nested past the renderer's depth instead of exhausting the stack (#448)", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts, onUnsupported: "warn" });
        const pres = new PuppeteerGen(PAGE, { renderer, createDocument });

        pres.addSlide().addImage({ data: `data:image/svg+xml;base64,${btoa(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1">${"<g>".repeat(300)}<rect width="1" height="1"/>${"</g>".repeat(300)}</svg>`)}`, x: 1, y: 1, w: 1, h: 1 });
        const document = await pres.toHtml();

        await renderer.render(document);
        expect(renderer.warnings()).toEqual(["slide 1: object 'Image 0': an SVG image's elements nest deeper than 256 levels; the picture is not drawn"]);
        await expect(renderer.render({ ...document, html: document.html.replace("</body>", `${"<div>".repeat(300)}</body>`) })).rejects.toThrow("invalid PdfLapas HTML: elements nest deeper than 256 levels");
    });

    test("stops parsing at the depth limit, so deep nesting costs milliseconds, not seconds (GHSA-mc68-2mrj-jwhm)", async () => {
        const renderer = await createPdfkitRenderer({ fonts: systemFonts, onUnsupported: "warn" });
        const pres = new PuppeteerGen(PAGE, { renderer, createDocument });

        // HTML elements inside <svg> leave SVG parsing, which is where parse5 is quadratic in depth (100k took 14 s)
        pres.addSlide().addImage({ data: `data:image/svg+xml;base64,${btoa(`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1 1">${"<div>".repeat(100_000)}</svg>`)}`, x: 1, y: 1, w: 1, h: 1 });
        const document = await pres.toHtml();
        let start = performance.now();

        await renderer.render(document);
        expect(performance.now() - start).toBeLessThan(500);
        expect(renderer.warnings()).toEqual(["slide 1: object 'Image 0': an SVG image's elements nest deeper than 256 levels; the picture is not drawn"]);
        const html = document.html.replace("</body>", `${"<div>".repeat(100_000)}</body>`);

        start = performance.now();
        await expect(renderer.render({ ...document, html })).rejects.toThrow("invalid PdfLapas HTML: elements nest deeper than 256 levels");
        expect(performance.now() - start).toBeLessThan(500);
    });

    test("runs a right-to-left paragraph's words from the right, and reports mixed directions", async () => {
        const renderer = await createPdfkitRenderer({ fonts: await loadSystemFonts({ match: /^LiberationSans-Regular\.ttf$/ }), onUnsupported: "warn" });
        const pdf = await render(renderer, pres => pres.addSlide().addText("שלום עולם", { x: 1, y: 1, w: 5, h: 1, fontFace: "Arial", align: "right", rtlMode: true }));

        expect(renderer.warnings()).toEqual([]);
        const [first, second] = texts(pageContents(pdf)[0] ?? "");

        expect(first?.x ?? 0).toBeGreaterThan(second?.x ?? 0);
        await render(renderer, pres => pres.addSlide().addText("PDF שלום", { x: 1, y: 1, w: 5, h: 1, fontFace: "Arial", rtlMode: true }));
        expect(renderer.warnings()).toEqual(["slide 1: object 'Text 0': left-to-right words in a right-to-left paragraph are ordered in the paragraph's direction"]);
    });

    test("synthesizes bold and italic where the face has neither, as the HTML says", async () => {
        const regularOnly = await loadSystemFonts({ match: /^LiberationSans-Regular\.ttf$/ });
        const renderer = await createPdfkitRenderer({ fonts: regularOnly });
        const pdf = await render(renderer, pres => pres.addSlide().addText("Bold", { x: 1, y: 1, w: 2, h: 1, fontFace: "Arial", fontSize: 18, bold: true, italic: true, margin: 0 }));
        const content = pageContents(pdf)[0] ?? "";

        // The LibreOffice target writes a 1/30 em stroke (24 px / 30) and a 0.3462535606 shear about the baseline.
        expect(content).toContain("0.8 w");
        expect(content).toContain("2 Tr");
        expect(content).toMatch(/1 0 -0\.34625\d* 1 [\d.]+ 0 cm/);
    });

    test("needs font bytes, and reads system fonts only under Node/Bun", async () => {
        await expect(createBrowserPdfkitRenderer({ fonts: [] })).rejects.toThrow("createPdfkitRenderer() requires at least one font");
        await expect(createBrowserPdfkitRenderer({ fonts: [new Uint8Array([1, 2, 3])] })).rejects.toThrow("Font #0 is not a TTF, OTF or TTC font file");
        await expect(browserLoadSystemFonts()).rejects.toThrow("only available under Node/Bun");
        await expect(createBrowserPdfkitRenderer({ fonts: systemFonts, onUnsupported: "ignore" as never })).rejects.toThrow('onUnsupported must be "error" or "warn"');
    });

    test("breaks lines where Chromium does: after spaces, not inside a minus before a digit", () => {
        const split = (text: string) => {
            const cuts = [0, ...breakOpportunities(text), text.length];

            return cuts.slice(1).map((end, index) => text.slice(cuts[index], end));
        };

        expect(split("Hello world. Again")).toEqual(["Hello ", "world. ", "Again"]);
        expect(split("a - b x-1 AA-22")).toEqual(["a ", "- ", "b ", "x-", "1 ", "AA-", "22"]);
        expect(split("foo  bar")).toEqual(["foo  ", "bar"]);
    });
});
