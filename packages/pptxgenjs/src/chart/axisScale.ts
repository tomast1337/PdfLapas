import { PPTX_DEFAULTS } from "#src/defaults";

/** A value axis's bounds and the step between its labels and gridlines. */
export type AxisScale = { minimum: number; maximum: number; interval: number };

// Floating point slack when a value is a whole number of intervals (0.9 / 0.1 is 9.000000000000002).
const STEP_EPSILON = 1e-9;
// The candidate intervals within each power of ten, smallest first.
const INTERVAL_FACTORS = [1, 2, 5] as const;

/** Rounds away the floating point noise of a multiple of a decimal interval (0.30000000000000004 is 0.3). */
function decimal(value: number): number {
    return Number(value.toPrecision(12));
}

/**
 * LibreOffice's automatic interval of a linear axis from minimum to maximum (PPTX_DEFAULTS.chart.automaticScale): the
 * smallest 1, 2 or 5 times a power of ten that splits the range into at most maxIncrements steps.
 */
export function automaticAxisInterval(minimum: number, maximum: number, maxIncrements: number = PPTX_DEFAULTS.chart.automaticScale.maxIncrements): number {
    const range = maximum - minimum;

    if (!(range > 0) || !Number.isFinite(range)) {
        return 1;
    }

    // range / 10^(floor(log10(range)) - 1) is 10 to 100, so ten times that power always fits in maxIncrements steps.
    const magnitude = 10 ** (Math.floor(Math.log10(range)) - 1);

    for (const power of [magnitude, magnitude * 10]) {
        for (const factor of INTERVAL_FACTORS) {
            const interval = decimal(factor * power);

            if (range / interval <= maxIncrements + STEP_EPSILON) {
                return interval;
            }
        }
    }

    return decimal(magnitude * 100);
}

/** The step between an axis's labels: automatic in at most maxIncrements steps, or fixed by the caller. */
export type AxisStep = { kind: "automatic"; maxIncrements: number } | { kind: "fixed"; interval: number };

/** The bound an option fixes on an axis whose other bound is automatic, or none. */
export type FixedAxisBound = { kind: "none" } | { kind: "minimum"; value: number } | { kind: "maximum"; value: number };

/**
 * LibreOffice's automatic scale of a linear axis whose data runs from low to high (PPTX_DEFAULTS.chart.automaticScale),
 * ported from chart2 ScaleAutomatism::calculateExplicitIncrementAndScaleForLinear:
 * 1. A fixed bound beyond the data takes the data's place (a fixed minimum above the data is also the maximum).
 * 2. A range at or below 0 is scaled negated, [-5, -2] as [2, 5] with its fixed bound on the other side.
 * 3. A value axis (towardZero) whose automatic minimum is above 0 drops it to 0 when it is under wideRatio of the
 *    maximum (or equal to it), else moves it half the range toward 0.
 * 4. A range still empty doubles an automatic maximum (0 becomes 1) or halves an automatic minimum (0 becomes -1).
 * 5. The interval is automaticAxisInterval's. Each automatic bound rounds out to whole intervals counted from the fixed
 *    bound (from 0 when none is fixed), then moves one interval further where the data comes within borderRatio of the
 *    range of it, unless it is 0. More than maxIncrements steps take the next interval in 1, 2, 5 x 10^k.
 * A scatter chart's x axis passes its data range, with no fixed bound and not toward 0. A renderer that knows the drawn
 * axis passes the steps its labels leave room for as maxIncrements (the target's valueAxisStepRoom, #585). A fixed step
 * (valAxisMajorUnit) is the interval, and the automatic bounds round out to it: 3..8 by 2 is 0..10 in LibreOffice
 * 26.2.6, -15..30 by 7 is -21..35 (#595).
 */
export function automaticAxisScale(low: number, high: number, fixed: FixedAxisBound, towardZero: boolean, step: AxisStep = { kind: "automatic", maxIncrements: PPTX_DEFAULTS.chart.automaticScale.maxIncrements }): AxisScale {
    const { borderRatio, wideRatio } = PPTX_DEFAULTS.chart.automaticScale;
    let autoMinimum = fixed.kind !== "minimum";
    let autoMaximum = fixed.kind !== "maximum";
    let sourceMinimum = fixed.kind === "minimum" ? fixed.value : low;
    let sourceMaximum = fixed.kind === "maximum" ? fixed.value : high;

    if (sourceMinimum > sourceMaximum) {
        if (autoMaximum) {
            sourceMaximum = sourceMinimum;
        } else {
            sourceMinimum = sourceMaximum;
        }
    }

    const negate = sourceMinimum < 0 && sourceMaximum <= 0;

    if (negate) {
        [sourceMinimum, sourceMaximum] = [-sourceMaximum, -sourceMinimum];
        [autoMinimum, autoMaximum] = [autoMaximum, autoMinimum];
    }

    let tempMinimum = sourceMinimum;
    let tempMaximum = sourceMaximum;

    if (towardZero && autoMinimum && tempMinimum > 0) {
        if (tempMinimum === tempMaximum || tempMinimum / tempMaximum < wideRatio) {
            tempMinimum = 0;
        } else {
            tempMinimum -= (tempMaximum - tempMinimum) / 2;
        }
    }

    if (tempMinimum === tempMaximum) {
        if (autoMaximum) {
            tempMaximum = tempMaximum === 0 ? 1 : tempMaximum * 2;
        } else {
            tempMinimum = tempMinimum === 0 ? -1 : tempMinimum / 2;
        }
    }

    // The intervals count from the fixed bound, so a fixed minimum of 7 draws 7, 12, 17 and not 10, 15.
    const base = !autoMinimum ? tempMinimum : !autoMaximum ? tempMaximum : 0;
    let interval = step.kind === "fixed" ? step.interval : automaticAxisInterval(tempMinimum, tempMaximum, step.maxIncrements);
    let scale: AxisScale = { minimum: low, maximum: high, interval };

    // Each pass at least doubles the interval, so three passes reach one that fits.
    for (let pass = 0; pass < 3; pass++) {
        let minimum = tempMinimum;
        let maximum = tempMaximum;

        if (autoMinimum) {
            minimum = decimal(base + Math.floor((minimum - base) / interval + STEP_EPSILON) * interval);

            if (minimum !== 0 && (maximum - sourceMinimum) / (maximum - minimum) > 1 - borderRatio) {
                minimum = decimal(minimum - interval);
            }
        }

        if (autoMaximum) {
            maximum = decimal(base + Math.ceil((maximum - base) / interval - STEP_EPSILON) * interval);

            if (maximum !== 0 && (sourceMaximum - minimum) / (maximum - minimum) > 1 - borderRatio) {
                maximum = decimal(maximum + interval);
            }
        }

        // `+ 0` turns a negated 0 into 0, so no label reads -0.
        scale = negate ? { minimum: -maximum + 0, maximum: -minimum + 0, interval } : { minimum, maximum, interval };

        if (step.kind === "fixed" || Math.floor((maximum - minimum) / interval + STEP_EPSILON) <= step.maxIncrements) {
            return scale;
        }

        // The next of 1, 2 and 5 x 10^k.
        interval = decimal(interval * (interval.toExponential(0).startsWith("2") ? 2.5 : 2));
    }

    return scale;
}

/** The bounds and step a caller set on a value axis; each one left out is automatic. */
export type FixedAxisScale = { minimum?: number; maximum?: number; majorUnit?: number };

/**
 * LibreOffice's scale of a 3D bar chart's value axis whose data runs from low to high (0 included): the automatic
 * interval in at most maxIncrements steps (from the axis's drawn length, see renderBar3DSvg), the automatic bounds
 * rounded out to whole intervals from the minimum, with no extra interval near the data, unlike a 2D axis (chart2
 * BarChart: no border expansion in 3D). Measured in LibreOffice 26.2.6: 3, 5, 4, 7 gives 0..7 by 1 on a tall axis and
 * 0..10 by 5 for 3, 5, 4, 8 on a short one; valAxisMajorUnit 2 gives 0..8, and 1..9 from a valAxisMinVal of 1.
 */
export function bar3DAxisScale(low: number, high: number, fixed: FixedAxisScale, maxIncrements: number): AxisScale {
    let interval = fixed.majorUnit ?? automaticAxisInterval(fixed.minimum ?? low, fixed.maximum ?? high, maxIncrements);
    let scale: AxisScale = { minimum: low, maximum: high, interval };

    // As in automaticAxisScale, a scale whose rounded bounds make too many steps takes the next interval.
    for (let pass = 0; pass < 3; pass++) {
        const minimum = fixed.minimum ?? Math.floor(low / interval + STEP_EPSILON) * interval;
        const maximum = fixed.maximum ?? minimum + Math.max(1, Math.ceil((high - minimum) / interval - STEP_EPSILON)) * interval;

        scale = { minimum: decimal(minimum), maximum: decimal(maximum), interval };

        if (fixed.majorUnit !== undefined || (maximum - minimum) / interval <= maxIncrements + STEP_EPSILON) {
            return scale;
        }

        interval = automaticAxisInterval(minimum, maximum, maxIncrements);
    }

    return scale;
}
