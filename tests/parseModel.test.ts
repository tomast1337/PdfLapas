import { describe, expect, test } from "bun:test";
import { pointsToPixels } from "@pdflapas/model";
import { PPTX_DEFAULTS } from "../packages/pptxgenjs/src/defaults";
import { InvalidModelError, parseModel } from "../packages/pptxgenjs/src/modelSchema";
import { createNodeDocument, PuppeteerGen } from "../packages/pptxgenjs/src/node";
import { modelToHtml } from "../packages/pptxgenjs/src/render/document";
import type { ComponentLookup } from "../packages/pptxgenjs/src/render/slide";
import { buildCrossRuntimeDeck } from "./crossRuntimeDeck";
import { populateParityFixture } from "./visual/fixture";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

function smallModel(): unknown {
    const pres = new PuppeteerGen();

    pres.addSlide().addText("Hi", { x: 0.5, y: 0.5, w: 2, h: 1, color: "FF0000", fill: { color: "00FF00" } });
    pres.addSlide().addImage({ data: PNG, x: 0, y: 0, w: 1, h: 1 });

    return JSON.parse(JSON.stringify(pres.getModel()));
}

type Json = Record<string | number, unknown>;

/** The InvalidModelError parseModel() throws for `model`. */
function invalid(model: unknown): InvalidModelError {
    try {
        parseModel(model);
    } catch (error) {
        if (error instanceof InvalidModelError) {
            return error;
        }

        throw error;
    }

    throw new Error("parseModel() accepted the model");
}

/** The InvalidModelError parseModel() throws for the small model with `value` written at `path` (its last key the field). */
function rejection(path: readonly (string | number)[], value: unknown): InvalidModelError {
    const model = smallModel();
    // `as`: smallModel() is getModel() JSON and every test path names objects and arrays it has.
    let parent = model as Json;

    for (const key of path.slice(0, -1)) {
        parent = parent[key] as Json;
    }

    parent[path[path.length - 1] ?? ""] = value;

    return invalid(model);
}

describe("parseModel()", () => {
    test("getModel() of every fixture deck round-trips through JSON and parseModel(), and modelToHtml() of it is toHtml()", async () => {
        const decks: Record<string, PuppeteerGen> = { fixture: new PuppeteerGen(), "cross-runtime": buildCrossRuntimeDeck(PuppeteerGen, PNG) };

        populateParityFixture(decks.fixture!);
        // Imported here: tests/visual/wasm/deck.ts installs a global jsdom `document` when loaded (see htmlBytes.test.ts).
        const hadDocument = "document" in globalThis;

        try {
            const { DECKS } = await import("./visual/wasm/decks");

            for (const [name, deck] of Object.entries(DECKS)) {
                decks[`wasm-${name}`] = deck.populate(new PuppeteerGen());
            }

            for (const [name, pres] of Object.entries(decks)) {
                const json = JSON.parse(JSON.stringify(pres.getModel()));

                expect({ name, model: parseModel(json) }).toEqual({ name, model: json });
                // `as`: the registry and resolver are private; toHtml() hands the same objects to modelToHtml().
                const { components, resolveImage } = pres as unknown as { components: ComponentLookup; resolveImage: (source: string) => Promise<string> };
                const html = (await modelToHtml(json, { components, resolveImage, createDocument: createNodeDocument })).html;

                expect({ name, same: html === (await pres.toHtml()).html }).toEqual({ name, same: true });
            }
        } finally {
            if (!hadDocument) {
                Reflect.deleteProperty(globalThis, "document");
            }
        }
    }, 120_000);

    test("rejects a color that would add CSS declarations", () => {
        const error = rejection(["slides", 0, "objects", 0, "text", "paragraphs", 0, "runs", 0, "color"], "red;background:url(x)");

        expect(error.message).toContain("model.slides[0].objects[0].text.paragraphs[0].runs[0].color: is not a CssColor");
        expect(rejection(["slides", 0, "background"], "#FF0000; color: red").issues).toEqual([{ path: "model.slides[0].background", message: expect.stringContaining("colorToCSS") }]);
    });

    test("rejects a page size that is NaN, negative, zero or infinite", () => {
        expect(rejection(["pageSize", "width"], Number.NaN).message).toContain("model.pageSize.width: ");
        expect(rejection(["pageSize", "height"], -1).message).toContain("model.pageSize.height: must be a finite number of inches above 0");
        expect(rejection(["pageSize", "height"], 0).issues[0]?.path).toBe("model.pageSize.height");
        expect(rejection(["pageSize", "width"], Number.POSITIVE_INFINITY).issues[0]?.path).toBe("model.pageSize.width");
    });

    test("rejects a wrong type, an unknown key and an unknown object kind", () => {
        expect(rejection(["slides", 0, "hidden"], "no").message).toContain("model.slides[0].hidden: Invalid input: expected boolean");
        expect(rejection(["slides", 0, "objects", 0, "style", "geometry", "x"], "10px").issues[0]?.path).toBe("model.slides[0].objects[0].style.geometry.x");
        expect(rejection(["slides", 0, "fontSize"], 12).message).toContain("model.slides[0]: Unrecognized key");
        expect(rejection(["slides", 0, "objects", 0, "kind"], "video").message).toContain("model.slides[0].objects[0].kind: is not a slide object");
        expect(() => parseModel("deck")).toThrow(InvalidModelError);
    });

    test("rejects a data image whose source is not a data URI, a link that runs script and a font face that ends its declaration", () => {
        expect(rejection(["slides", 1, "objects", 0, "image", "source"], "https://example.com/a.png").message).toContain("model.slides[1].objects[0].image.source: is not a data URI");
        expect(rejection(["slides", 1, "objects", 0, "image", "source"], "data:text/html;base64,PHNjcmlwdD4=").issues[0]?.path).toBe("model.slides[1].objects[0].image.source");
        expect(rejection(["slides", 0, "objects", 0, "text", "paragraphs", 0, "runs", 0, "link"], { href: "javascript:alert(1)" }).message).toContain("runs[0].link.href: is not a supported link");
        expect(rejection(["slides", 0, "objects", 0, "text", "paragraphs", 0, "runs", 0, "fontFace"], 'Arial";}\nbody{color:red').message).toContain("runs[0].fontFace: ");
    });

    test("rejects a shadow blur or glow size over the API limit and accepts it at the limit", () => {
        const shadowBlur = pointsToPixels(PPTX_DEFAULTS.limits.shadowBlurPt);
        const glowSize = pointsToPixels(PPTX_DEFAULTS.limits.glowSizePt);
        // The NormalizedShadow both effects use; a glow carries its size (the API's `glow.size`) in `blur`.
        const effect = (blur: number) => ({ visible: true, color: "#000000", blur, offsetX: 0, offsetY: 0, inset: false });

        function withShadow(blur: number): unknown {
            const model = smallModel() as { slides: { objects: { style: { shadow?: unknown } }[] }[] };

            model.slides[0]!.objects[0]!.style.shadow = effect(blur);

            return model;
        }

        function withGlow(blur: number): unknown {
            const model = smallModel() as { slides: { objects: { text: { paragraphs: { runs: { glow?: unknown }[] }[] } }[] }[] };

            model.slides[0]!.objects[0]!.text.paragraphs[0]!.runs[0]!.glow = effect(blur);

            return model;
        }

        expect(() => parseModel(withShadow(shadowBlur))).not.toThrow();
        expect(() => parseModel(withGlow(glowSize))).not.toThrow();

        const shadow = invalid(withShadow(shadowBlur + 1));

        expect(shadow.message).toContain(`model.slides[0].objects[0].style.shadow.blur: shadow.blur is ${shadowBlur + 1} px, above the API limit of 100 pt (${shadowBlur} px)`);
        const glow = invalid(withGlow(glowSize + 1));

        expect(glow.message).toContain(`model.slides[0].objects[0].text.paragraphs[0].runs[0].glow.blur: glow.blur is ${glowSize + 1} px, above the API limit of 150 pt (${glowSize} px)`);
        // checkedSize refuses a negative blur too, so a foreign model cannot ask for one either.
        expect(invalid(withShadow(-1)).issues[0]?.path).toBe("model.slides[0].objects[0].style.shadow.blur");
        expect(invalid(withGlow(-1)).issues[0]?.path).toBe("model.slides[0].objects[0].text.paragraphs[0].runs[0].glow.blur");
    });

    test("rejects word letter-spacing that is not one finite value per word of the run", () => {
        const spacing = ["slides", 0, "objects", 0, "text", "paragraphs", 0, "runs", 0, "wordCharacterSpacing"] as const;

        // "Hi" is one word.
        expect(rejection(spacing, [0.1, 0.2]).message).toContain("runs[0].wordCharacterSpacing: ");
        expect(rejection(spacing, []).issues[0]?.path).toBe("model.slides[0].objects[0].text.paragraphs[0].runs[0].wordCharacterSpacing");
        expect(rejection(spacing, [Number.NaN]).issues[0]?.path).toBe("model.slides[0].objects[0].text.paragraphs[0].runs[0].wordCharacterSpacing[0]");
        // A word cut at its kern pairs: its pieces cover its characters.
        expect(rejection(spacing, [[{ characters: 1, spacing: 0.1 }]]).message).toContain("must cover the word");
        expect(rejection(spacing, [[{ characters: 0, spacing: 0.1 }]]).message).toContain("runs[0].wordCharacterSpacing[0]");
        // "Hi" has no stretch of spaces, and space letter-spacing comes only with word letter-spacing.
        const spaces = ["slides", 0, "objects", 0, "text", "paragraphs", 0, "runs", 0, "spaceCharacterSpacing"] as const;

        expect(rejection(spaces, [0.1]).message).toContain("runs[0].spaceCharacterSpacing: ");
        expect(rejection(spaces, []).message).toContain("comes with wordCharacterSpacing");
    });

    test("writes a word cut at its kern pairs as a word span per piece, side by side", async () => {
        const model = smallModel() as { slides: { objects: { text: { paragraphs: { runs: Json[] }[] } }[] }[] };
        const run = model.slides[0]?.objects[0]?.text.paragraphs[0]?.runs[0];

        if (!run) {
            throw new Error("the small model has no run");
        }

        run.wordCharacterSpacing = [
            [
                { characters: 1, spacing: -0.5 },
                { characters: 1, spacing: 0.25 },
            ],
        ];
        const { html } = await modelToHtml(parseModel(model), { createDocument: createNodeDocument });

        expect(html).toContain('<span style="letter-spacing: -0.5px;">H</span><span style="letter-spacing: 0.25px;">i</span>');
    });

    test("modelToHtml() checks its input; the error names the path", async () => {
        const model = smallModel() as { slides: { color: string }[] };

        model.slides[0]!.color = "red;background:url(x)";
        await expect(modelToHtml(model)).rejects.toThrow("model.slides[0].color: is not a CssColor");
    });
});
