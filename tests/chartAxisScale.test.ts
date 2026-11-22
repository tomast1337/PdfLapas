import { describe, expect, test } from "bun:test";
import type PptxGenJS from "pptxgenjs";
import { automaticAxisInterval, automaticAxisScale } from "../packages/pptxgenjs/src/chart/axisScale";
import { normalizeChart } from "../packages/pptxgenjs/src/normalize/chart";

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

    test("matches LibreOffice 26.2's bar, line and area value axes, which reach to 0 (#547)", () => {
        // Measured in its PDF export of synthetic PptxGenJS 4.0.1 charts: chart, options, series values, then the drawn axis.
        const cases: [PptxGenJS.CHART_NAME, PptxGenJS.IChartOpts, number[][], number, number, number][] = [
            ["bar", {}, [[1, 4, 2]], 0, 4.5, 0.5],
            ["bar", {}, [[3, 7, 5]], 0, 8, 1],
            ["bar", {}, [[3, 10, 5]], 0, 12, 2],
            ["bar", {}, [[3, 11, 5]], 0, 12, 2],
            ["bar", {}, [[5, 12, 8]], 0, 14, 2],
            ["bar", {}, [[16, 9, 4]], 0, 18, 2],
            ["bar", {}, [[20, 9, 4]], 0, 25, 5],
            ["bar", {}, [[48, 20, 30]], 0, 60, 10],
            ["bar", {}, [[95, 20, 30]], 0, 100, 10],
            ["bar", {}, [[100, 20, 30]], 0, 120, 20],
            ["bar", {}, [[100, 150, 160]], 0, 180, 20],
            ["bar", {}, [[0.2, 0.9, 0.5]], 0, 1, 0.1],
            ["bar", {}, [[1, 2.5, 2]], 0, 3, 0.5],
            ["bar", {}, [[-3, 8, 5]], -4, 10, 2],
            ["bar", {}, [[-12, 5, 7]], -15, 10, 5],
            ["bar", {}, [[-7, -2, -4]], -8, 0, 1],
            [
                "bar",
                { barGrouping: "stacked" },
                [
                    [5, 4, 3],
                    [7, 3, 2],
                ],
                0,
                14,
                2,
            ],
            [
                "bar",
                { barGrouping: "stacked", barDir: "bar" },
                [
                    [8, 6, 2],
                    [8, 5, 3],
                ],
                0,
                18,
                2,
            ],
            ["bar", { barDir: "bar" }, [[48, 20, 30]], 0, 60, 10],
            ["line", {}, [[5, 12, 8]], 0, 14, 2],
            ["line", {}, [[100, 150, 160]], 0, 180, 20],
            ["line", {}, [[-12, 5, 7]], -15, 10, 5],
            ["area", {}, [[16, 9, 4]], 0, 18, 2],
            [
                "area",
                { barGrouping: "stacked" },
                [
                    [5, 4, 3],
                    [7, 3, 2],
                ],
                0,
                14,
                2,
            ],
        ];

        for (const [type, options, values, minimum, maximum, interval] of cases) {
            const data = values.map((series, index) => ({ name: `S${index}`, labels: ["A", "B", "C"], values: series }));
            const chart = normalizeChart(type, data, options);

            expect([chart.valueAxisMinimum, chart.valueAxisMaximum, automaticAxisInterval(chart.valueAxisMinimum, chart.valueAxisMaximum)]).toEqual([minimum, maximum, interval]);
        }
    });
});
