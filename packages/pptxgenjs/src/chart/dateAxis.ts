/**
 * A date category axis (`<c:dateAx>`), which PptxGenJS 4.0.1 writes for catLabelFormatCode: its categories are Excel
 * date serials, and Office lays them out on a time line of one slot per base time unit from the first date to the
 * last, leaving a slot without a date empty (#580).
 */

export type ChartTimeUnit = "days" | "months" | "years";

type CalendarDate = { year: number; month: number; day: number };

/** A date-format token: a field of the date, or literal text. */
type DateToken = { kind: "year"; digits: 2 | 4 } | { kind: "month"; width: 1 | 2 | 3 | 4 | 5 } | { kind: "day"; width: 1 | 2 | 3 | 4 } | { kind: "text"; text: string };

const MONTH_NAMES = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"] as const;
const WEEKDAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"] as const;
const MS_PER_DAY = 86_400_000;
// Excel's 1900 date system: serial 1 is 1900-01-01 and serial 60 the 1900-02-29 Lotus kept, so from serial 61 on a
// serial counts days from 1899-12-30.
const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
const EXCEL_FAKE_LEAP_DAY = 60;
// Characters Excel prints as they are in a format code; any other letter is a field or needs quoting.
const LITERAL_CHARACTERS = new Set([" ", "-", "/", ".", ",", ":", "(", ")", "'"]);

// shortcut: a time line longer than this many slots (13 years of days) is rejected, as each slot is a category of the
// drawn axis; draw it as a continuous time axis when a deck needs a longer one.
export const DATE_AXIS_SLOT_LIMIT = 5000;

/** The tokens of an Excel date format code, or why it is not one PdfLapas formats. */
function dateFormatTokens(formatCode: string): { kind: "tokens"; tokens: DateToken[] } | { kind: "unsupported" } {
    const tokens: DateToken[] = [];
    let index = 0;

    while (index < formatCode.length) {
        const character = formatCode.charAt(index);
        const lower = character.toLowerCase();

        if (character === '"') {
            const end = formatCode.indexOf('"', index + 1);

            if (end < 0) {
                return { kind: "unsupported" };
            }

            tokens.push({ kind: "text", text: formatCode.slice(index + 1, end) });
            index = end + 1;

            continue;
        }

        if (character === "\\") {
            tokens.push({ kind: "text", text: formatCode.charAt(index + 1) });
            index += 2;

            continue;
        }

        if (lower === "y" || lower === "m" || lower === "d") {
            let run = 1;

            while (formatCode.charAt(index + run).toLowerCase() === lower) {
                run++;
            }

            if (lower === "y") {
                tokens.push({ kind: "year", digits: run <= 2 ? 2 : 4 });
            } else if (lower === "m") {
                tokens.push({ kind: "month", width: Math.min(run, 5) as 1 | 2 | 3 | 4 | 5 });
            } else {
                tokens.push({ kind: "day", width: Math.min(run, 4) as 1 | 2 | 3 | 4 });
            }

            index += run;

            continue;
        }

        // Hours, seconds, AM/PM, sections, colors and digit placeholders are not date fields this formats.
        if (!LITERAL_CHARACTERS.has(character)) {
            return { kind: "unsupported" };
        }

        tokens.push({ kind: "text", text: character });
        index++;
    }

    return tokens.some(token => token.kind !== "text") ? { kind: "tokens", tokens } : { kind: "unsupported" };
}

/** Whether a catLabelFormatCode is an Excel date format of years, months, days and literal text. */
export function isSupportedChartDateFormat(formatCode: string): boolean {
    return dateFormatTokens(formatCode).kind === "tokens";
}

function serialToDate(serial: number): CalendarDate {
    const whole = Math.floor(serial);

    if (whole === EXCEL_FAKE_LEAP_DAY) {
        return { year: 1900, month: 1, day: 29 };
    }

    const date = new Date(EXCEL_EPOCH_MS + (whole < EXCEL_FAKE_LEAP_DAY ? whole + 1 : whole) * MS_PER_DAY);

    return { year: date.getUTCFullYear(), month: date.getUTCMonth(), day: date.getUTCDate() };
}

function dateToSerial(date: CalendarDate): number {
    return (Date.UTC(date.year, date.month, date.day) - EXCEL_EPOCH_MS) / MS_PER_DAY;
}

/** An Excel date serial in a date format code (isSupportedChartDateFormat); the serial itself in any other. */
export function formatChartDate(serial: number, formatCode: string): string {
    const format = dateFormatTokens(formatCode);

    if (format.kind === "unsupported") {
        return String(serial);
    }

    const date = serialToDate(serial);
    let text = "";
    const tokens = format.tokens;

    for (const token of tokens) {
        if (token.kind === "text") {
            text += token.text;
        } else if (token.kind === "year") {
            text += token.digits === 2 ? String(date.year % 100).padStart(2, "0") : String(date.year);
        } else if (token.kind === "month") {
            const name = MONTH_NAMES[date.month] ?? "";

            text += token.width === 1 ? String(date.month + 1) : token.width === 2 ? String(date.month + 1).padStart(2, "0") : token.width === 3 ? name.slice(0, 3) : token.width === 4 ? name : name.charAt(0);
        } else {
            const weekday = WEEKDAY_NAMES[new Date(Date.UTC(date.year, date.month, date.day)).getUTCDay()] ?? "";

            text += token.width === 1 ? String(date.day) : token.width === 2 ? String(date.day).padStart(2, "0") : token.width === 3 ? weekday.slice(0, 3) : weekday;
        }
    }

    return text;
}

/** A date's position on a time line of `unit` slots: its serial day, month count or year. */
function unitIndex(date: CalendarDate, unit: ChartTimeUnit): number {
    return unit === "days" ? Date.UTC(date.year, date.month, date.day) / MS_PER_DAY : unit === "months" ? date.year * 12 + date.month : date.year;
}

/** `date` moved by `count` units, its day kept within the month it lands in. */
function addUnits(date: CalendarDate, count: number, unit: ChartTimeUnit): CalendarDate {
    if (unit === "days") {
        const moved = new Date(Date.UTC(date.year, date.month, date.day + count));

        return { year: moved.getUTCFullYear(), month: moved.getUTCMonth(), day: moved.getUTCDate() };
    }

    const months = date.year * 12 + date.month + (unit === "months" ? count : count * 12);
    const year = Math.floor(months / 12);
    const month = months - year * 12;
    const monthDays = new Date(Date.UTC(year, month + 1, 0)).getUTCDate();

    return { year, month, day: Math.min(date.day, monthDays) };
}

export type DateAxisOptions = {
    /** Each category's Excel date serial. */
    serials: readonly number[];
    /** catAxisBaseTimeUnit; absent: years when every date shares its month and day, months when its day, else days. */
    baseUnit?: ChartTimeUnit;
    /** catAxisMajorUnit and catAxisMajorTimeUnit (the base unit when absent): a label every that many units. */
    majorUnit?: number;
    majorTimeUnit?: ChartTimeUnit;
    /** catAxisMinVal and catAxisMaxVal: the time line's first and last dates as serials. */
    minimum?: number;
    maximum?: number;
};

export type DateAxisLayout =
    | {
          kind: "slots";
          /** Each slot's date serial, null for a slot between two major units, which has no label. */
          labelSerials: (number | null)[];
          /**
           * Each category's slot, null when it is outside catAxisMinVal..catAxisMaxVal. shortcut: LibreOffice draws a line
           * to such a point as far as the plot's edge, PdfLapas not at all; draw it when a deck needs it.
           */
          categorySlots: (number | null)[];
      }
    | { kind: "shared-slot"; unit: ChartTimeUnit }
    | { kind: "too-many-slots"; count: number };

/**
 * The time line LibreOffice 26.2.6 draws for a date axis (its PDF export of synthetic PptxGenJS 4.0.1 line and column
 * charts, #580): a slot per base unit from the first date (or catAxisMinVal) to the last (or catAxisMaxVal), each
 * category in its date's slot, and a label on every slot, or every catAxisMajorUnit from the first.
 */
export function dateAxisLayout(options: DateAxisOptions): DateAxisLayout {
    const dates = options.serials.map(serialToDate);
    let baseUnit = options.baseUnit;

    if (!baseUnit) {
        const [first] = dates;
        const sameDay = first !== undefined && dates.every(date => date.day === first.day);

        baseUnit = sameDay && dates.every(date => date.month === first.month) ? "years" : sameDay ? "months" : "days";
    }

    let firstSerial = Infinity;
    let lastSerial = -Infinity;

    for (const serial of options.serials) {
        firstSerial = Math.min(firstSerial, serial);
        lastSerial = Math.max(lastSerial, serial);
    }

    const start = serialToDate(options.minimum ?? firstSerial);
    const startIndex = unitIndex(start, baseUnit);
    const slotCount = unitIndex(serialToDate(options.maximum ?? lastSerial), baseUnit) - startIndex + 1;

    if (slotCount > DATE_AXIS_SLOT_LIMIT) {
        return { kind: "too-many-slots", count: slotCount };
    }

    const categorySlots: (number | null)[] = [];
    const taken = new Set<number>();

    for (const date of dates) {
        const slot = unitIndex(date, baseUnit) - startIndex;

        if (taken.has(slot)) {
            return { kind: "shared-slot", unit: baseUnit };
        }

        taken.add(slot);
        categorySlots.push(slot >= 0 && slot < slotCount ? slot : null);
    }

    const labelled = new Set<number>();

    if (options.majorUnit === undefined) {
        for (let slot = 0; slot < slotCount; slot++) {
            labelled.add(slot);
        }
    } else {
        // Each step moves at least one day, so the slots grow and the loop ends past the last one.
        for (let step = 0; ; step++) {
            const slot = unitIndex(addUnits(start, step * options.majorUnit, options.majorTimeUnit ?? baseUnit), baseUnit) - startIndex;

            if (slot >= slotCount) {
                break;
            }

            labelled.add(slot);
        }
    }

    const labelSerials: (number | null)[] = [];

    for (let slot = 0; slot < slotCount; slot++) {
        labelSerials.push(labelled.has(slot) ? dateToSerial(addUnits(start, slot, baseUnit)) : null);
    }

    return { kind: "slots", labelSerials, categorySlots };
}
