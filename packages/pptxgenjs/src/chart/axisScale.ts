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
export function automaticAxisInterval(minimum: number, maximum: number): number {
    const range = maximum - minimum;

    if (!(range > 0) || !Number.isFinite(range)) {
        return 1;
    }

    // range / 10^(floor(log10(range)) - 1) is 10 to 100, so ten times that power always fits in maxIncrements steps.
    const magnitude = 10 ** (Math.floor(Math.log10(range)) - 1);

    for (const power of [magnitude, magnitude * 10]) {
        for (const factor of INTERVAL_FACTORS) {
            const interval = decimal(factor * power);

            if (range / interval <= PPTX_DEFAULTS.chart.automaticScale.maxIncrements + STEP_EPSILON) {
                return interval;
            }
        }
    }

    return decimal(magnitude * 100);
}

/**
 * LibreOffice's automatic scale of a linear axis whose data runs from low to high, as it draws a scatter chart's x axis
 * (it does not reach to zero): the automatic interval, the bounds rounded out to whole intervals, then each bound one
 * interval further where the data comes within borderRatio of the range of it, except a minimum of 0; a bound that
 * makes too many steps takes the next interval.
 */
export function automaticAxisScale(low: number, high: number): AxisScale {
    const { maxIncrements, borderRatio } = PPTX_DEFAULTS.chart.automaticScale;
    let interval = automaticAxisInterval(low, high);
    let scale: AxisScale = { minimum: low, maximum: high, interval };

    // Each pass at most doubles to quintuples the interval, so three passes reach one that fits.
    for (let pass = 0; pass < 3; pass++) {
        let minimum = Math.floor(low / interval + STEP_EPSILON) * interval;
        let maximum = Math.ceil(high / interval - STEP_EPSILON) * interval;

        if (maximum <= minimum) {
            maximum = minimum + interval;
        }

        const border = (maximum - minimum) * borderRatio;

        if (low !== 0 && low - minimum <= border + STEP_EPSILON) {
            minimum -= interval;
        }

        if (maximum - high <= border + STEP_EPSILON) {
            maximum += interval;
        }

        scale = { minimum: decimal(minimum), maximum: decimal(maximum), interval };

        if ((maximum - minimum) / interval <= maxIncrements + STEP_EPSILON) {
            return scale;
        }

        interval = automaticAxisInterval(minimum, maximum);
    }

    return scale;
}
