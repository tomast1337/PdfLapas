import { describe, expect, test } from "bun:test";
import type PptxGenJS from "pptxgenjs";
import { automaticAxisInterval, automaticAxisScale, bar3DAxisScale, type FixedAxisScale } from "../packages/pptxgenjs/src/chart/axisScale";
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
            expect(automaticAxisScale(low, high, { kind: "none" }, false)).toEqual({ minimum, maximum, interval });
        }
    });

    test("takes the smallest 1, 2 or 5 x 10^k interval with at most ten steps", () => {
        expect(automaticAxisInterval(0, 10)).toBe(1);
        expect(automaticAxisInterval(0, 11)).toBe(2);
        expect(automaticAxisInterval(0, 7)).toBe(1);
        expect(automaticAxisInterval(0, 0.7)).toBe(0.1);
        expect(automaticAxisInterval(-30, 45)).toBe(10);
        expect(automaticAxisInterval(0, 1e6)).toBe(1e5);
        // No range keeps a unit step; ScaleAutomatism doubles the maximum of a single value, 5 scaling as 5..10 (drawn 4..10).
        expect(automaticAxisInterval(5, 5)).toBe(1);
        expect(automaticAxisScale(5, 5, { kind: "none" }, false)).toEqual({ minimum: 4, maximum: 10, interval: 1 });
    });

    test("takes no more steps than maxIncrements, the room the drawn axis leaves its labels (#585)", () => {
        // 3, -2, 4, 7: PowerPoint's 11 steps draw -3..8 by 1, LibreOffice's 10 -4..8 by 2, a short axis -5..10 by 5.
        expect(automaticAxisScale(-2, 7, { kind: "none" }, true, { kind: "automatic", maxIncrements: 11 })).toEqual({ minimum: -3, maximum: 8, interval: 1 });
        expect(automaticAxisScale(-2, 7, { kind: "none" }, true)).toEqual({ minimum: -4, maximum: 8, interval: 2 });
        expect(automaticAxisScale(-2, 7, { kind: "none" }, true, { kind: "automatic", maxIncrements: 4 })).toEqual({ minimum: -5, maximum: 10, interval: 5 });
        expect(automaticAxisScale(3, 7, { kind: "none" }, true, { kind: "automatic", maxIncrements: 2 })).toEqual({ minimum: 0, maximum: 10, interval: 5 });
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

    test("matches LibreOffice 26.2's value axes beside a fixed bound and on narrow data (#556)", () => {
        // Measured in its PDF export of synthetic PptxGenJS 4.0.1 charts (4.4 x 4.1in, Arial): chart, options, series
        // values, then the drawn axis. The renderer counts the labels from the minimum, so 7..52 by 5 reads 7, 12, 17.
        const cases: [PptxGenJS.CHART_NAME, PptxGenJS.IChartOpts, number[][], number, number, number][] = [
            ["bar", { valAxisMinVal: 5 }, [[10, 48, 30]], 5, 55, 5],
            ["bar", { valAxisMinVal: 20 }, [[3, 48, 30]], 20, 50, 5],
            ["bar", { valAxisMinVal: 7 }, [[10, 48, 30]], 7, 52, 5],
            ["bar", { valAxisMinVal: -10 }, [[3, 12, 8]], -10, 15, 5],
            ["bar", { valAxisMinVal: 50 }, [[60, 100, 80]], 50, 110, 10],
            ["line", { valAxisMinVal: 2 }, [[1, 4, 3]], 2, 4.5, 0.5],
            ["bar", { valAxisMinVal: 60 }, [[10, 48, 30]], 60, 120, 10],
            ["line", { valAxisMinVal: 5 }, [[10, 48, 30]], 5, 55, 5],
            ["bar", { valAxisMinVal: 5, barDir: "bar" }, [[10, 48, 30]], 5, 55, 5],
            ["bar", { valAxisMaxVal: 30 }, [[3, 48, 20]], 0, 30, 5],
            ["bar", { valAxisMaxVal: 100 }, [[10, 48, 30]], 0, 100, 10],
            ["bar", { valAxisMaxVal: 50 }, [[-12, 30, 20]], -20, 50, 10],
            ["bar", { valAxisMaxVal: 10 }, [[-3, 8, 5]], -4, 10, 2],
            ["bar", { valAxisMaxVal: 40 }, [[-48, -10, -20]], -60, 40, 10],
            ["bar", { valAxisMaxVal: -5 }, [[-48, -10, -20]], -55, -5, 5],
            ["bar", { valAxisMaxVal: 3 }, [[1, 4, 2]], 0, 3, 0.5],
            ["bar", { valAxisMaxVal: 45 }, [[-20, 48, 10]], -25, 45, 10],
            ["bar", { valAxisMaxVal: -5 }, [[-48, -2, -20]], -55, -5, 5],
            ["area", { valAxisMaxVal: 30 }, [[3, 48, 20]], 0, 30, 5],
            [
                "bar",
                { valAxisMinVal: 5, barGrouping: "stacked" },
                [
                    [5, 30, 10],
                    [5, 18, 20],
                ],
                5,
                55,
                5,
            ],
            ["bar", { valAxisMaxVal: 50 }, [[44, 48, 46]], 41, 50, 1],
            ["bar", {}, [[150, 155, 160]], 144, 162, 2],
            ["bar", { valAxisMinVal: -5 }, [[-48, -10, -20]], -5, 0, 0.5],
            ["bar", { valAxisMaxVal: 0 }, [[3, 48, 20]], -1, 0, 0.1],
            ["bar", { valAxisMinVal: -60 }, [[-48, -10, -20]], -60, 0, 10],
            ["line", {}, [[150, 155, 160]], 144, 162, 2],
            ["bar", { valAxisMaxVal: -50 }, [[-48, -10, -20]], -100, -50, 5],
            ["bar", { valAxisMinVal: 0 }, [[3, 48, 20]], 0, 60, 10],
            ["bar", {}, [[-150, -155, -160]], -162, -144, 2],
            ["bar", { valAxisMinVal: 140 }, [[150, 155, 160]], 140, 165, 5],
            ["bar", { valAxisMinVal: 5 }, [[1, 4, 2]], 5, 10, 0.5],
            ["bar", { valAxisMaxVal: -2 }, [[-4, -1, -2]], -4.5, -2, 0.5],
        ];

        for (const [type, options, values, minimum, maximum, interval] of cases) {
            const data = values.map((series, index) => ({ name: `S${index}`, labels: ["A", "B", "C"], values: series }));
            const chart = normalizeChart(type, data, options);

            expect([chart.valueAxisMinimum, chart.valueAxisMaximum, automaticAxisInterval(chart.valueAxisMinimum, chart.valueAxisMaximum)]).toEqual([minimum, maximum, interval]);
        }
    });

    test("matches LibreOffice 26.2's bar3D value axes, with no step added near the data (#541)", () => {
        // Measured in its PDF export of synthetic PptxGenJS 4.0.1 bar3D charts: data low and high, the caller's bounds
        // and unit, the step count the axis's length fits (8 on a 9 x 4.8in chart in 14pt labels), then the drawn axis.
        const cases: [number, number, FixedAxisScale, number, number, number, number][] = [
            [0, 7, {}, 8, 0, 7, 1],
            [0, 8, {}, 8, 0, 8, 1],
            [0, 9, {}, 8, 0, 10, 2],
            [0, 16, {}, 8, 0, 16, 2],
            [0, 18, {}, 8, 0, 20, 5],
            [0, 20, {}, 8, 0, 20, 5],
            [0, 95, {}, 8, 0, 100, 20],
            [0, 1.4, {}, 8, 0, 1.4, 0.2],
            [0, 350, {}, 8, 0, 350, 50],
            [0, 4.5, {}, 8, 0, 5, 1],
            [-3, 7, {}, 8, -4, 8, 2],
            [0, 11, {}, 8, 0, 12, 2],
            [0, 7, { majorUnit: 2 }, 8, 0, 8, 2],
            [0, 7, { majorUnit: 2.5 }, 8, 0, 7.5, 2.5],
            [0, 9, { majorUnit: 2, minimum: 1 }, 8, 1, 9, 2],
            [0, 7, { maximum: 10 }, 8, 0, 10, 2],
            [0, 7, { minimum: 2, maximum: 9 }, 8, 2, 9, 1],
            [0, 7, { maximum: 20, majorUnit: 5 }, 8, 0, 20, 5],
            // Shorter charts (9 x 2.2in and 9 x 1.4in) fit fewer steps.
            [0, 8, {}, 6, 0, 8, 2],
            [0, 8, {}, 3, 0, 10, 5],
        ];

        for (const [low, high, fixed, maxIncrements, minimum, maximum, interval] of cases) {
            expect(bar3DAxisScale(low, high, fixed, maxIncrements)).toEqual({ minimum, maximum, interval });
        }
    });
});
