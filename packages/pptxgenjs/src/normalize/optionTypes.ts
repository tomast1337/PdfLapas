/** A JavaScript type an option accepts; "array" is an Array or a typed array, "object" any other non-null object. */
type JsType = "string" | "number" | "boolean" | "object" | "array";

/** The types an option accepts, and the types of the fields or entries it holds. */
export type OptionType = {
    readonly types: readonly JsType[];
    /** Each field's name and type, as entries listed once here rather than on every check. */
    readonly fields?: readonly (readonly [string, OptionType])[];
    readonly items?: OptionType;
    /** Whether an entry may be missing (null or undefined); otherwise each one is checked against `items`. */
    readonly holes?: boolean;
};

const STRING: OptionType = { types: ["string"] };
const NUMBER: OptionType = { types: ["number"] };
const OBJECT: OptionType = { types: ["object"] };

const fields = (record: Readonly<Record<string, OptionType>>) => Object.entries(record);

const TYPE_WORDS: Readonly<Record<JsType, string>> = { string: "a string", number: "a number", boolean: "true or false", object: "an object", array: "an array" };

/**
 * The options normalization reads with a string method, writes into an attribute, iterates or reads a field of, so a
 * value of another type would end in a JavaScript-engine TypeError that names no option (#674). PptxGenJS declares
 * these types; it reads `null`, and a falsy value where it takes an object, as absent (its checks are truthiness checks).
 */
const OPTIONS: OptionType = {
    types: ["object"],
    fields: fields({
        fontFace: STRING,
        lang: STRING,
        objectName: STRING,
        shapeName: STRING,
        altText: STRING,
        fit: STRING,
        vert: STRING,
        textDirection: STRING,
        margin: { types: ["number", "array"] },
        // A string is the deprecated line color, which PptxGenJS still reads.
        line: { types: ["object", "string"] },
        shadow: OBJECT,
        tabStops: { types: ["array"], items: OBJECT },
        points: { types: ["array"], items: OBJECT },
        hyperlink: { types: ["object"], fields: fields({ url: STRING, tooltip: STRING }) },
        bullet: { types: ["boolean", "object"], fields: fields({ style: STRING }) },
        sizing: { types: ["object"], fields: fields({ type: STRING }) },
        // One object for every side, or [top, right, bottom, left]; a missing side is none, as in PptxGenJS.
        border: { types: ["object", "array"], items: OBJECT, holes: true },
        lineDataSymbol: STRING,
        catLabelFormatCode: STRING,
        dataLabelFormatCode: STRING,
        dataTableFormatCode: STRING,
        serLabelFormatCode: STRING,
        valAxisLabelFormatCode: STRING,
        valLabelFormatCode: STRING,
        valAxisMinVal: NUMBER,
        valAxisMaxVal: NUMBER,
        valAxes: { types: ["array"], items: { types: ["object"], fields: fields({ valAxisMinVal: NUMBER, valAxisMaxVal: NUMBER, valAxisLabelFormatCode: STRING }) } },
        catAxes: { types: ["array"], items: { types: ["object"], fields: fields({ catLabelFormatCode: STRING }) } },
    }),
};

/** A chart series: its labels, values and bubble sizes are iterated. */
const SERIES: OptionType = { types: ["object"], fields: fields({ labels: { types: ["array"] }, values: { types: ["array"] }, sizes: { types: ["array"] } }) };

/**
 * Throws a TypeError naming the option when `value` (an options object of `method`, or a chart series with
 * `series: true`) holds an option of a type PptxGenJS does not declare for it.
 */
export function checkOptionTypes(value: unknown, method: string, series = false): void {
    checkType(value, series ? SERIES : OPTIONS, method, "");
}

function checkType(value: unknown, type: OptionType, method: string, name: string): void {
    // PptxGenJS reads these options with truthiness checks: null is one left out, and so is any falsy value (0, "",
    // false) where it takes an object, as for `hyperlink: 0` (#456).
    if (value === undefined || value === null || (!value && type.types.includes("object"))) {
        return;
    }

    const kind = typeof value;
    // A typed array is read as an array of numbers (chart values).
    const actual: JsType | "other" = Array.isArray(value) || ArrayBuffer.isView(value) ? "array" : kind === "string" || kind === "number" || kind === "boolean" || kind === "object" ? kind : "other";

    if (actual === "other" || !type.types.includes(actual)) {
        const got = typeof value === "string" ? JSON.stringify(value) : actual === "other" ? `a ${kind}` : actual === "array" || actual === "object" ? TYPE_WORDS[actual] : String(value);
        const words: string[] = [];

        for (const accepted of type.types) {
            words.push(TYPE_WORDS[accepted]);
        }

        throw new TypeError(`${method}: ${name === "" ? "the options" : `option \`${name}\``} must be ${words.join(" or ")}, got ${got}`);
    }

    if (Array.isArray(value)) {
        const items = type.items;

        if (items) {
            for (const [index, item] of value.entries()) {
                if (!type.holes && (item === undefined || item === null)) {
                    throw new TypeError(`${method}: option \`${name}[${index}]\` must be ${TYPE_WORDS[items.types[0] ?? "object"]}, got ${String(item)}`);
                }

                checkType(item, items, method, `${name}[${index}]`);
            }
        }

        return;
    }

    if (type.fields && typeof value === "object") {
        for (const [key, field] of type.fields) {
            // `as`: a non-null, non-array object, read by key.
            checkType((value as Record<string, unknown>)[key], field, method, name === "" ? key : `${name}.${key}`);
        }
    }
}
