import { describe, expect, test } from "bun:test";
import type PptxGenJS from "pptxgenjs";
import type { FontMetricsLookup, GlyphAdvances } from "../packages/pptxgenjs/src/defaults";
import { fallbackChain, fallbackSegments } from "../packages/pptxgenjs/src/normalize/fallback";
import { normalizeText } from "../packages/pptxgenjs/src/normalize/text";

// Which of the measured characters each face maps, as their cmaps do (PMingLiU as SimSun, ⅐ in the Arial PowerPoint drew).
const COVERAGE: Readonly<Record<string, string>> = {
    Arial: " ax⅐©™☺1",
    "Times New Roman": " ax⅐↉",
    "Cambria Math": "⇒①",
    "Segoe UI Symbol": "☆✓⭐😀①⇒🇯🇵",
    "Segoe UI Emoji": "✓⭐😀❤⇒🇯🇵☺1\uFE0F\u20E3",
    "MS Gothic": "是あア働込ｱ〜，☆①⇒❤",
    "Yu Gothic": "够是あア働込ｱ〜，↉☆①⇒",
    "Microsoft JhengHei": "够这是ㄅ，あ☆①",
    "Microsoft YaHei": "够这是ㄅ䶶，",
    "Malgun Gothic": "한국어是，☆①⇒",
    PMingLiU: "够这是ㄅ，",
    SimSun: "够这是ㄅ䶶，",
};
const advances = (characters: string): GlyphAdvances => ({ unitsPerEm: 1000, advances: new Map([...characters].map(character => [character.codePointAt(0) ?? 0, 1000])) });
const UNKNOWN = { kind: "unknown" } as const;
const FONTS: FontMetricsLookup = Object.assign(() => UNKNOWN, { glyphAdvances: (face: string) => (Object.hasOwn(COVERAGE, face) ? ({ kind: "advances", advances: advances(COVERAGE[face] ?? "") } as const) : UNKNOWN) });

/** The faces a renderer draws `text` in, in order: per character, the first of the run's face and its stretch's chain that has it. */
function drawn(text: string, fontFace = "Arial"): [string, string][] {
    const faces: [string, string][] = [];

    for (const { text: stretch, fallbackFaces } of fallbackSegments(text, fontFace, false, false, "en-US", "powerpoint", FONTS)) {
        for (const character of stretch) {
            const face = [fontFace, ...fallbackFaces].find(candidate => COVERAGE[candidate]?.includes(character)) ?? "none";
            const last = faces.at(-1);

            if (last?.[1] === face) {
                last[0] += character;
            } else {
                faces.push([character, face]);
            }
        }
    }

    return faces;
}

describe("the PowerPoint target's fallback faces", () => {
    test("writes Arial's and Calibri's linked East Asian faces after its symbol faces, and Times New Roman's for other faces", () => {
        const linked = ["MS Gothic", "Yu Gothic", "Microsoft JhengHei", "Microsoft YaHei", "Malgun Gothic"];

        expect(fallbackChain("Arial", "en-US", "powerpoint")).toEqual(["Cambria Math", "Segoe UI Symbol", ...linked]);
        expect(fallbackChain("CALIBRI", "en-US", "powerpoint")).toEqual(["Arial", "Cambria Math", "Segoe UI Symbol", ...linked]);
        expect(fallbackChain("Times New Roman", "en-US", "powerpoint")).toEqual(["Arial", "Cambria Math", "Segoe UI Symbol", "Yu Gothic", "PMingLiU", "SimSun", "Malgun Gothic"]);
        expect(fallbackChain("constructor", "en-US", "powerpoint")).toEqual(fallbackChain("Times New Roman", "en-US", "powerpoint"));
    });

    // Measured against PowerPoint's PDF export of a synthetic deck (#403): "x这是x x够x xあアx x한글x" in ten faces.
    test("links Arial's East Asian faces to the other sans-serif faces, Yu Gothic first to the monospace ones", () => {
        const linked = ["MS Gothic", "Yu Gothic", "Microsoft JhengHei", "Microsoft YaHei", "Malgun Gothic"];
        const symbols = ["Arial", "Cambria Math", "Segoe UI Symbol"];

        for (const face of ["Aptos", "Segoe UI", "Verdana", "Tahoma"]) {
            expect(fallbackChain(face, "en-US", "powerpoint")).toEqual([...symbols, ...linked]);
        }

        for (const face of ["Courier New", "Consolas"]) {
            expect(fallbackChain(face, "en-US", "powerpoint")).toEqual([...symbols, "Yu Gothic", "Microsoft JhengHei", "Microsoft YaHei", "Malgun Gothic"]);
        }

        expect(fallbackChain("Georgia", "en-US", "powerpoint")).toEqual(fallbackChain("Times New Roman", "en-US", "powerpoint"));
        expect(fallbackChain("Cambria", "en-US", "powerpoint")).toEqual(fallbackChain("Times New Roman", "en-US", "powerpoint"));
    });

    // Measured against PowerPoint's PDF export of a synthetic deck: each line's faces in Arial (Calibri and Arial bold
    // alike) and in Times New Roman, at 14pt.
    test("keeps an East Asian character in the face of the one before it while that face has it, as PowerPoint does", () => {
        expect(drawn("够是够是")).toEqual([["够是够是", "Yu Gothic"]]);
        expect(drawn("这是这是")).toEqual([["这是这是", "Microsoft JhengHei"]]);
        expect(drawn("这够")).toEqual([["这够", "Microsoft JhengHei"]]);
        expect(drawn("够这")).toEqual([
            ["够", "Yu Gothic"],
            ["这", "Microsoft JhengHei"],
        ]);
        expect(drawn("ㄅ是")).toEqual([["ㄅ是", "Microsoft JhengHei"]]);
        expect(drawn("这，是")).toEqual([["这，是", "Microsoft JhengHei"]]);
        expect(drawn("䶶是")).toEqual([["䶶是", "Microsoft YaHei"]]);
        expect(drawn("한국어是")).toEqual([["한국어是", "Malgun Gothic"]]);
        expect(drawn("あア働込是")).toEqual([["あア働込是", "MS Gothic"]]);
        expect(drawn("ｱ〜是")).toEqual([["ｱ〜是", "MS Gothic"]]);
        // A space, a character of the run's face or another character starts over.
        expect(drawn("这 是")).toEqual([
            ["这", "Microsoft JhengHei"],
            [" ", "Arial"],
            ["是", "MS Gothic"],
        ]);
        expect(drawn("这a是")).toEqual([
            ["这", "Microsoft JhengHei"],
            ["a", "Arial"],
            ["是", "MS Gothic"],
        ]);
        expect(drawn("⅐是↉是")).toEqual([
            ["⅐", "Arial"],
            ["是", "MS Gothic"],
            ["↉", "Yu Gothic"],
            ["是", "MS Gothic"],
        ]);
        expect(drawn("够是", "Times New Roman")).toEqual([["够是", "Yu Gothic"]]);
        expect(drawn("这是 是", "Times New Roman")).toEqual([
            ["这是", "PMingLiU"],
            [" ", "Times New Roman"],
            ["是", "Yu Gothic"],
        ]);
        expect(drawn("䶶是", "Times New Roman")).toEqual([["䶶是", "SimSun"]]);
    });

    // Measured in the same export: "x<c>x yo" in Arial at 18pt and lines of emoji, drawn from Segoe UI Emoji at its
    // 1.373em advance; ☆ and ✓ (not Extended_Pictographic) stay in Segoe UI Symbol, ⇒ and ① in Cambria Math. ❤,
    // whose default presentation is text, follows the chain (#403).
    test("draws an emoji the run's face lacks from Segoe UI Emoji when its default presentation is emoji, other symbols from the chain", () => {
        expect(drawn("x😀x")).toEqual([
            ["x", "Arial"],
            ["😀", "Segoe UI Emoji"],
            ["x", "Arial"],
        ]);
        expect(drawn("⭐❤🇯🇵")).toEqual([
            ["⭐", "Segoe UI Emoji"],
            ["❤", "MS Gothic"],
            ["🇯🇵", "Segoe UI Emoji"],
        ]);
        expect(drawn("x☆✓⇒①x")).toEqual([
            ["x", "Arial"],
            ["☆✓", "Segoe UI Symbol"],
            ["⇒①", "Cambria Math"],
            ["x", "Arial"],
        ]);
    });

    // Measured in a second export (#403): ☺ ™ ↔ ♥ ♠ ‼ are laid out at the face's advances where it has them, and at
    // Segoe UI Symbol's where it lacks them, never at Segoe UI Emoji's, and drawn black; so is 1️⃣.
    test("keeps an emoji whose default presentation is text, and a keycap, out of Segoe UI Emoji", () => {
        expect(drawn("x☺©™x")).toEqual([["x☺©™x", "Arial"]]);
        expect(drawn("x☺x", "Times New Roman")).toEqual([
            ["x", "Times New Roman"],
            ["☺", "Arial"],
            ["x", "Times New Roman"],
        ]);
        expect(fallbackSegments("x1\uFE0F\u20E3x", "Arial", false, false, "en-US", "powerpoint", FONTS)).toEqual([{ text: "x1\uFE0F\u20E3x", fallbackFaces: fallbackChain("Arial", "en-US", "powerpoint") }]);
    });

    // In its PDF export of a synthetic deck (#588), a character followed by U+FE0F is the face's own text glyph where
    // the face has it (a keycap's "1", "☺️" in Arial), and a colour emoji where it lacks it ("☺️" in Calibri, "❤️"):
    // U+FE0E asks every renderer for that text glyph. LibreOffice was not measured and keeps U+FE0F.
    test("writes U+FE0F after a character the face has as U+FE0E on the PowerPoint target only", () => {
        const texts = (target: "powerpoint" | "libreoffice", fontFace: string, fonts = FONTS) =>
            normalizeText([{ text: "x1\uFE0F\u20E3x \u263A\uFE0F 2\u20E3 \u2764\uFE0F" }], { fontFace, fontSize: 14 }, target, fonts)
                .paragraphs.flatMap(paragraph => paragraph.runs.map(run => run.text))
                .join("");
        const unknown: FontMetricsLookup = Object.assign(() => UNKNOWN, { glyphAdvances: () => UNKNOWN });

        expect(texts("powerpoint", "Arial")).toBe("x1\uFE0E\u20E3x \u263A\uFE0E 2\u20E3 \u2764\uFE0F");
        expect(texts("powerpoint", "Times New Roman")).toBe("x1\uFE0F\u20E3x \u263A\uFE0F 2\u20E3 \u2764\uFE0F");
        // A face whose characters are not known: only a keycap's base is taken as the face's.
        expect(texts("powerpoint", "Arial", unknown)).toBe("x1\uFE0E\u20E3x \u263A\uFE0F 2\u20E3 \u2764\uFE0F");
        expect(texts("libreoffice", "Arial")).toBe("x1\uFE0F\u20E3x \u263A\uFE0F 2\u20E3 \u2764\uFE0F");
    });

    test("leaves the text whole when a face's characters are unknown, and on the LibreOffice target", () => {
        const unknown: FontMetricsLookup = Object.assign(() => UNKNOWN, { glyphAdvances: () => UNKNOWN });

        expect(fallbackSegments("够这", "Arial", false, false, "en-US", "powerpoint", unknown)).toEqual([{ text: "够这", fallbackFaces: fallbackChain("Arial", "en-US", "powerpoint") }]);
        expect(fallbackSegments("够这", "Arial", false, false, "en-US", "libreoffice", FONTS)).toEqual([{ text: "够这", fallbackFaces: fallbackChain("Arial", "en-US", "libreoffice") }]);
    });

    test("cuts a run into one run a stretch, the first keeping its soft break", () => {
        const runs = (input: PptxGenJS.TextProps[]) => normalizeText(input, { fontFace: "Arial", fontSize: 14 }, "powerpoint", FONTS).paragraphs.flatMap(paragraph => paragraph.runs.map(run => [run.text, run.fallbackFaces?.[0], run.softBreakBefore]));

        expect(runs([{ text: "a" }, { text: "够是", options: { softBreakBefore: true } }])).toEqual([
            ["a", "Cambria Math", false],
            ["够", "Cambria Math", true],
            ["是", "Yu Gothic", false],
        ]);
    });

    // Measured in PowerPoint's PDF export of a synthetic deck (#403): "这"|"是" in two runs of identical options is all
    // Microsoft JhengHei; a run in another colour or weight, or after a soft break, starts 是 over in MS Gothic.
    test("carries the held East Asian face into the next run when their options are identical", () => {
        const firstFaces = (input: PptxGenJS.TextProps[]) => normalizeText(input, { fontFace: "Arial", fontSize: 14 }, "powerpoint", FONTS).paragraphs.flatMap(paragraph => paragraph.runs.map(run => [run.text, run.fallbackFaces?.[0]]));

        expect(firstFaces([{ text: "这" }, { text: "是" }, { text: "是" }])).toEqual([
            ["这", "Cambria Math"],
            ["是", "Microsoft JhengHei"],
            ["是", "Microsoft JhengHei"],
        ]);
        expect(firstFaces([{ text: "这" }, { text: "是", options: { color: "C00000" } }])).toEqual([
            ["这", "Cambria Math"],
            ["是", "Cambria Math"],
        ]);
        expect(firstFaces([{ text: "这" }, { text: "是", options: { bold: true } }])).toEqual([
            ["这", "Cambria Math"],
            ["是", "Cambria Math"],
        ]);
        expect(firstFaces([{ text: "这" }, { text: "是", options: { softBreakBefore: true } }])).toEqual([
            ["这", "Cambria Math"],
            ["是", "Cambria Math"],
        ]);
    });
});

// Measured in LibreOffice 26.2's PDF export in the visual image (RENDERING_TARGETS.libreoffice.genericFallback, #289).
describe("the LibreOffice target's fallback faces", () => {
    test("follows the run face's generic family: sans-serif, serif or monospace", () => {
        expect(fallbackChain("Arial", "en-US", "libreoffice")).toEqual(["DejaVu Sans", "Noto Sans CJK SC"]);
        expect(fallbackChain("Segoe UI", "en-US", "libreoffice")).toEqual(["DejaVu Sans", "Noto Sans CJK SC"]);
        expect(fallbackChain("Times New Roman", "en-US", "libreoffice")).toEqual(["DejaVu Serif", "DejaVu Sans", "Noto Serif CJK SC"]);
        expect(fallbackChain("CAMBRIA", "en-US", "libreoffice")).toEqual(["DejaVu Serif", "DejaVu Sans", "Noto Serif CJK SC"]);
        expect(fallbackChain("Courier New", "en-US", "libreoffice")).toEqual(["DejaVu Sans Mono", "DejaVu Sans", "Noto Sans Mono CJK SC"]);
        // A face of the chain is left out of it.
        expect(fallbackChain("DejaVu Serif", "en-US", "libreoffice")).toEqual(["DejaVu Sans", "Noto Serif CJK SC"]);
    });

    test("takes the East Asian face's region from the run's language, its tag then its primary subtag", () => {
        const region = (language: string) => fallbackChain("Arial", language, "libreoffice").at(-1);

        expect([region("zh-CN"), region("zh-Hans"), region("fr-FR")]).toEqual(["Noto Sans CJK SC", "Noto Sans CJK SC", "Noto Sans CJK SC"]);
        expect([region("ZH-tw"), region("zh-Hant-TW")]).toEqual(["Noto Sans CJK TC", "Noto Sans CJK TC"]);
        expect([region("zh"), region("zh-HK"), region("zh-MO"), region("zh-SG"), region("zh-Hant"), region("yue-HK")]).toEqual(Array(6).fill("Noto Sans CJK HK"));
        expect([region("ja"), region("ja-JP"), region("ko"), region("ko-KR"), region("ko-KP")]).toEqual(["Noto Sans CJK JP", "Noto Sans CJK JP", "Noto Sans CJK KR", "Noto Sans CJK KR", "Noto Sans CJK KR"]);
        // A tag named like an Object.prototype member is just another language.
        expect(region("constructor")).toBe("Noto Sans CJK SC");
        // The HK face is the sans-serif one in every family.
        expect([fallbackChain("Times New Roman", "zh-HK", "libreoffice").at(-1), fallbackChain("Consolas", "zh", "libreoffice").at(-1), fallbackChain("Georgia", "zh-TW", "libreoffice").at(-1)]).toEqual([
            "Noto Sans CJK HK",
            "Noto Sans CJK HK",
            "Noto Serif CJK TC",
        ]);
    });

    test("writes the chain and draws U+2011 as U+2010 on every run", () => {
        const runs = normalizeText([{ text: "lorem‑ipsum 1ᵉʳ" }, { text: "这", options: { fontFace: "Georgia", lang: "ja-JP" } }], { fontFace: "Arial", fontSize: 14 }, "libreoffice").paragraphs.flatMap(paragraph =>
            paragraph.runs.map(run => [run.text, run.fallbackFaces, run.nonBreakingHyphen]),
        );

        expect(runs).toEqual([
            ["lorem‑ipsum 1ᵉʳ", ["DejaVu Sans", "Noto Sans CJK SC"], "hyphen"],
            ["这", ["DejaVu Serif", "DejaVu Sans", "Noto Serif CJK JP"], undefined],
        ]);
    });

    // ❤ ✔ ⚠ ☀ ♠, which DejaVu Sans has, are drawn from Noto Color Emoji in Arial, Calibri and Times New Roman.
    test("cuts the run around each emoji, whose stretch lists Noto Color Emoji first, and looks up no face", () => {
        const unknown: FontMetricsLookup = Object.assign(() => UNKNOWN, {
            glyphAdvances: () => {
                throw new Error("looked up");
            },
        });
        const chain = ["DejaVu Sans", "Noto Sans CJK SC"];
        const emoji = ["Noto Color Emoji", ...chain];

        expect(fallbackSegments("a❤✔b 😀👍🏽 c", "Arial", false, false, "en-US", "libreoffice", unknown)).toEqual([
            { text: "a", fallbackFaces: chain },
            { text: "❤✔", fallbackFaces: emoji },
            { text: "b ", fallbackFaces: chain },
            { text: "😀👍🏽", fallbackFaces: emoji },
            { text: " c", fallbackFaces: chain },
        ]);
        expect(fallbackSegments("", "Arial", false, false, "en-US", "libreoffice", unknown)).toEqual([{ text: "", fallbackFaces: chain }]);
    });
});
