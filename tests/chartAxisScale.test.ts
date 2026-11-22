import { describe, expect, test } from "bun:test";
import { automaticAxisInterval, automaticAxisScale } from "../packages/pptxgenjs/src/chart/axisScale";

describe("automatic axis scale", () => {
    test("matches LibreOffice 26.2's scatter x axes (#519)", () => {
        // Measured in its PDF export of synthetic PptxGenJS 4.0.1 scatter charts: data range, then the drawn axis.
        const cases: [number, number, number, number, number][] = [
            [0, 7.6, 0, 8, 1],
            [0.2, 0.9, 0.1, 1, 0.1],
            [100, 160, 90, 170, 10],
            [0, 48, 0, 60, 10],
            [0, 15, 0, 16, 2],
            [0, 3, 0, 3.5, 0.5],
        ];

        for (const [low, high, minimum, maximum, interval] of cases) {
            expect(automaticAxisScale(low, high)).toEqual({ minimum, maximum, interval });
        }
    });

    test("takes the smallest 1, 2 or 5 x 10^k interval with at most ten steps", () => {
        expect(automaticAxisInterval(0, 10)).toBe(1);
        expect(automaticAxisInterval(0, 11)).toBe(2);
        expect(automaticAxisInterval(0, 7)).toBe(1);
        expect(automaticAxisInterval(0, 0.7)).toBe(0.1);
        expect(automaticAxisInterval(-30, 45)).toBe(10);
        expect(automaticAxisInterval(0, 1e6)).toBe(1e5);
        // A single value or no range keeps a unit step.
        expect(automaticAxisInterval(5, 5)).toBe(1);
        expect(automaticAxisScale(5, 5)).toEqual({ minimum: 4, maximum: 6, interval: 1 });
    });
});
