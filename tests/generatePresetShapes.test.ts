import { describe, expect, test } from "bun:test";
import { parsePresetShapes } from "../scripts/generate-preset-shapes";

const source = (shape: string) => `<presetShapeDefinitions>${shape}</presetShapeDefinitions>`;

describe("generate-preset-shapes", () => {
    test("reads a complete synthetic preset", () => {
        const parsed = parsePresetShapes(source(`<wave><gdLst><gd name="a" fmla="val 1"/></gdLst><pathLst><path w="2"><moveTo><pt x="0" y="a"/></moveTo><arcTo wR="1" hR="1" stAng="0" swAng="5400000"/><close/></path></pathLst></wave>`), ["wave"]);

        expect(parsed).toEqual({
            kind: "parsed",
            definitions: {
                wave: {
                    adjustments: [],
                    guides: [["a", "val 1"]],
                    paths: [
                        {
                            w: "2",
                            commands: [["M", "0", "a"], ["A", "1", "1", "0", "5400000"], ["Z"]],
                        },
                    ],
                },
            },
            textRects: {},
        });
    });

    test("names the shape, element and attribute of every missing required attribute instead of writing null", () => {
        const parsed = parsePresetShapes(source(`<wave><gdLst><gd name="a"/></gdLst><pathLst><path><lnTo><pt y="1"/></lnTo><arcTo wR="1" hR="1" stAng="0"/></path></pathLst></wave>`), ["wave"]);

        expect(parsed).toEqual({
            kind: "invalid",
            problems: [`wave: <pt> has no "x" attribute`, `wave: <arcTo> has no "swAng" attribute`, `wave: <gd> has no "fmla" attribute`],
        });
    });

    test("reports a missing guide attribute read for a text rectangle", () => {
        const parsed = parsePresetShapes(source(`<other><gdLst><gd fmla="val 1"/></gdLst><rect l="a" t="t" r="r" b="b"/></other>`), []);

        expect(parsed).toEqual({ kind: "invalid", problems: [`other: <gd> has no "name" attribute`] });
    });
});
