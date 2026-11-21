import { readFileSync, writeFileSync } from "node:fs";
import { JSDOM } from "jsdom";
import pptxgen from "pptxgenjs";

// Presets PdfLapas draws by hand on purpose, each group with its reason. Every other preset PptxGenJS 4.0.1 can
// write (its ShapeType values) is generated from presetShapeDefinitions.xml.
const HAND_DRAWN: readonly { readonly why: string; readonly names: readonly string[] }[] = [
    { why: "drawn as an HTML box, ellipse or line, not a path; roundRect reads rectRadius", names: ["rect", "roundRect", "ellipse", "line", "lineInv"] },
    { why: "not a preset: the path is the caller's points", names: ["custGeom"] },
    { why: "PptxGenJS writes their adjustments from angleRange and arcThicknessRatio", names: ["arc", "pie", "chord", "blockArc"] },
    { why: "drawn as an ellipse, as before: the generated four-arc path scores 0.0055 against 0.0054 in its WASM crop (#153)", names: ["flowChartConnector"] },
    { why: "drawn by hand: the generated geometry scores 0.0001 worse in its wide and WASM crops (#153)", names: ["actionButtonInformation"] },
    { why: "drawn by hand: the generated geometry scores 0.0001 worse in its tall crop, both backends (#153)", names: ["leftRightRibbon"] },
];

export type ParsedPresetShapes = { readonly kind: "parsed"; readonly definitions: Record<string, unknown>; readonly textRects: Record<string, unknown> } | { readonly kind: "invalid"; readonly problems: readonly string[] };

// Reads the paths and guides of each preset in `presets`, and every text rectangle, from presetShapeDefinitions.xml.
// Each problem (a missing preset, point or required attribute, an unknown path command) is collected, so one run
// names them all and nothing incomplete reaches generatedPresetShapes.ts.
export function parsePresetShapes(xml: string, presets: readonly string[]): ParsedPresetShapes {
    const document = new JSDOM(xml, { contentType: "text/xml" }).window.document;
    const problems: string[] = [];
    // The "" stand-in for a missing attribute never leaves this function: any problem makes the result "invalid".
    const required = (shape: string, element: Element, attribute: string) => {
        const value = element.getAttribute(attribute);

        if (value === null) {
            problems.push(`${shape}: <${element.localName}> has no "${attribute}" attribute`);

            return "";
        }

        return value;
    };
    const point = (shape: string, element: Element) => {
        const value = element.querySelector(":scope > pt");

        if (!value) {
            problems.push(`${shape}: <${element.localName}> has no <pt>`);

            return [];
        }

        return [required(shape, value, "x"), required(shape, value, "y")];
    };
    const command = (shape: string, element: Element): string[] => {
        const name = element.localName;

        if (name === "close") {
            return ["Z"];
        }

        if (name === "moveTo") {
            return ["M", ...point(shape, element)];
        }

        if (name === "lnTo") {
            return ["L", ...point(shape, element)];
        }

        if (name === "arcTo") {
            return ["A", required(shape, element, "wR"), required(shape, element, "hR"), required(shape, element, "stAng"), required(shape, element, "swAng")];
        }

        const points: string[] = [];

        for (const value of element.querySelectorAll(":scope > pt")) {
            points.push(required(shape, value, "x"), required(shape, value, "y"));
        }

        if (name === "quadBezTo") {
            return ["Q", ...points];
        }

        if (name === "cubicBezTo") {
            return ["C", ...points];
        }

        problems.push(`${shape}: unsupported path command <${name}>`);

        return [];
    };
    const guides = (shape: string, list: Element | null) => [...(list?.children ?? [])].map(element => [required(shape, element, "name"), required(shape, element, "fmla")] as const);

    const definitions: Record<string, unknown> = {};

    for (const name of presets) {
        const xmlName = name === "folderCorner" ? "foldedCorner" : name;
        const shape = [...document.documentElement.children].find(element => element.localName === xmlName);

        if (!shape) {
            problems.push(`${name}: no <${xmlName}> preset in the source`);
            continue;
        }

        const paths = [];

        for (const path of shape.querySelectorAll(":scope > pathLst > path")) {
            paths.push({
                ...(path.getAttribute("w") ? { w: path.getAttribute("w") } : {}),
                ...(path.getAttribute("h") ? { h: path.getAttribute("h") } : {}),
                ...(path.getAttribute("fill") ? { fill: path.getAttribute("fill") } : {}),
                ...(path.getAttribute("stroke") ? { stroke: path.getAttribute("stroke") } : {}),
                commands: [...path.children].map(element => command(xmlName, element)),
            });
        }

        definitions[name] = {
            adjustments: guides(xmlName, shape.querySelector(":scope > avLst")),
            guides: guides(xmlName, shape.querySelector(":scope > gdLst")),
            paths,
        };
    }

    // The text rectangle (<rect>) of every preset whose rectangle is not the whole box, keyed by DrawingML name, with
    // the adjustments and only the guides the rectangle reads.
    const TEXT_RECT_EDGES = ["l", "t", "r", "b"] as const;
    const textRects: Record<string, unknown> = {};

    for (const shape of document.documentElement.children) {
        const rect = shape.querySelector(":scope > rect");

        if (!rect) {
            continue;
        }

        const edges = TEXT_RECT_EDGES.map(edge => rect.getAttribute(edge) ?? edge);

        if (edges.join() === TEXT_RECT_EDGES.join()) {
            continue;
        }

        // Walk back from the rectangle: a guide is kept when a later kept guide reads it. Some presets redefine a
        // name (gear9's a1), and a use reads the closest definition above it.
        const needed = new Set<string>(edges);
        const kept: (readonly [string, string])[] = [];

        for (const guide of guides(shape.localName, shape.querySelector(":scope > gdLst")).reverse()) {
            const [name, formula] = guide;

            if (!needed.delete(name)) {
                continue;
            }

            for (const token of formula.split(/\s+/).slice(1)) {
                needed.add(token);
            }

            kept.push(guide);
        }

        textRects[shape.localName] = { adjustments: guides(shape.localName, shape.querySelector(":scope > avLst")), guides: kept.reverse(), rect: edges };
    }

    if (problems.length) {
        return { kind: "invalid", problems };
    }

    return { kind: "parsed", definitions, textRects };
}

if (import.meta.main) {
    const source = process.argv[2];

    if (!source) {
        throw new Error("Usage: bun scripts/generate-preset-shapes.ts <presetShapeDefinitions.xml>");
    }

    const WRITABLE: readonly string[] = Object.values(new pptxgen().ShapeType);
    const handDrawn = new Set(HAND_DRAWN.flatMap(group => group.names));
    const unknown = [...handDrawn].filter(name => !WRITABLE.includes(name));

    if (unknown.length) {
        throw new Error(`HAND_DRAWN names that PptxGenJS 4.0.1 cannot write: ${unknown.join(", ")}`);
    }

    const PRESETS = WRITABLE.filter(name => !handDrawn.has(name)).sort();
    const parsed = parsePresetShapes(readFileSync(source, "utf8"), PRESETS);

    if (parsed.kind === "invalid") {
        throw new Error(`${source} does not hold every preset PdfLapas generates; use the ECMA-376 presetShapeDefinitions.xml or fix these:\n${parsed.problems.join("\n")}`);
    }

    const { definitions, textRects } = parsed;
    const output =
        `// Generated from ECMA-376 presetShapeDefinitions.xml by scripts/generate-preset-shapes.ts.\n` +
        `// Keep this data declarative; rendering behavior belongs in normalize/generatedPreset.ts.\n` +
        `export type GeneratedPresetDefinition = {\n` +
        `    readonly adjustments: readonly (readonly [string, string])[];\n` +
        `    readonly guides: readonly (readonly [string, string])[];\n` +
        `    readonly paths: readonly {\n` +
        `        readonly w?: string;\n` +
        `        readonly h?: string;\n` +
        `        readonly fill?: string;\n` +
        `        readonly stroke?: string;\n` +
        `        readonly commands: readonly (readonly string[])[];\n` +
        `    }[];\n` +
        `};\n` +
        `export const GENERATED_PRESET_NAMES = ${JSON.stringify(PRESETS)} as const;\n` +
        `export type GeneratedPresetName = (typeof GENERATED_PRESET_NAMES)[number];\n` +
        `export const GENERATED_PRESET_SHAPES: Readonly<Record<GeneratedPresetName, GeneratedPresetDefinition>> = ${JSON.stringify(definitions)};\n` +
        `type GeneratedTextRect = {\n` +
        `    readonly adjustments: readonly (readonly [string, string])[];\n` +
        `    readonly guides: readonly (readonly [string, string])[];\n` +
        `    readonly rect: readonly [string, string, string, string];\n` +
        `};\n` +
        `export const GENERATED_TEXT_RECTS: Readonly<Record<string, GeneratedTextRect>> = ${JSON.stringify(textRects)};\n`;

    writeFileSync(new URL("../packages/pptxgenjs/src/generatedPresetShapes.ts", import.meta.url), output);
    console.log(`Generated ${Object.keys(definitions).length} preset definitions and ${Object.keys(textRects).length} text rectangles.`);
}
