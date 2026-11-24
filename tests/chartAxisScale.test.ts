import { describe, expect, test } from "bun:test";
import type PptxGenJS from "pptxgenjs";
import type { NormalizedChartAxisPadding } from "@pdflapas/model";
import { automaticAxisInterval, automaticAxisScale, bar3DAxisScale, type FixedAxisScale, paddedAxisScale } from "../packages/pptxgenjs/src/chart/axisScale";
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

describe("PowerPoint's padded automatic scale", () => {
    const column: NormalizedChartAxisPadding = { ratio: 0.05, maxSteps: 10, dropsHalfRange: true, bubbleRoom: 0, count: "bounds" };
    const scatter = { ...column, dropsHalfRange: false };
    const scale = (low: number, high: number, padding: NormalizedChartAxisPadding, steps = 100): string => {
        const { minimum, maximum, interval } = paddedAxisScale(low, high, { kind: "none" }, padding, { kind: "room", steps });

        return `${minimum}..${maximum}/${interval}`;
    };

    test("matches PowerPoint's PDF export of synthetic column, line and secondary value axes (#589, #619)", () => {
        // Data low and high, then PowerPoint's axis, on axes long enough for every step the padded range allows.
        const cases: [number, number, string][] = [
            // The minimum drops to 0 under 5/6 of the maximum, else half the range toward 0.
            [83, 100, "0..120/20"],
            [83.4, 100, "75..105/5"],
            [88, 100, "82..102/2"],
            [95, 100, "92..101/1"],
            [99, 100, "98.4..100.2/0.2"],
            [950, 1000, "920..1010/10"],
            [1000, 1200, "900..1250/50"],
            [5000, 5600, "4700..5700/100"],
            [100, 100, "0..120/20"],
            [-5, -5, "-6..0/1"],
            [-100, -83, "-120..0/20"],
            [-100, -95, "-101..-92/1"],
            [0.85, 1, "0.75..1.05/0.05"],
            [12, 13, "11.4..13.2/0.2"],
            // A bound that is not 0 pads by 5% of the range, and the padded range takes at most 10 steps.
            [0, 9.5, "0..10/1"],
            [0, 9.6, "0..12/2"],
            [0, 19, "0..20/2"],
            [0, 19.1, "0..25/5"],
            [0, 47.6, "0..50/5"],
            [0, 47.7, "0..60/10"],
            [0, 10, "0..12/2"],
            [-0.2, 10, "-2..12/2"],
            [-2, 7.9, "-4..10/2"],
            [-2, 7.4, "-4..8/2"],
            [-2, 7, "-3..8/1"],
            [-10, -0.1, "-12..0/2"],
            [-9.5, 0, "-10..0/1"],
            [-3, 3, "-4..4/1"],
            [-0.5, 9.5, "-2..10/2"],
            [6, 10, "0..12/2"],
            [3, 7.4, "0..8/1"],
            [0.0012, 0.0097, "0..0.012/0.002"],
        ];

        for (const [low, high, expected] of cases) {
            expect(`${low}..${high} ${scale(low, high, column)}`).toBe(`${low}..${high} ${expected}`);
        }
    });

    test("keeps a scatter axis's minimum far from 0 and pads it, and pads a bubble chart's by its room (#616, #620)", () => {
        expect(scale(950, 1000, scatter)).toBe("940..1010/10");
        expect(scale(85, 100, scatter)).toBe("84..102/2");
        expect(scale(5000, 5600, scatter)).toBe("4900..5700/100");
        expect(scale(-100, -95, scatter)).toBe("-101..-94/1");
        // 95..100 on a 4.2in wide plot, whose room leaves about four steps.
        expect(scale(95, 100, { ...scatter, count: "range" }, 4.5)).toBe("94..102/2");
        // Bubbles: y 3..8 0..10 by 1, x 0..100 -50..150 by 50 on a narrow plot, y 20..90 0..120 by 20.
        const bubble = { ...scatter, bubbleRoom: 0.15 };

        expect(scale(3, 8, bubble)).toBe("0..10/1");
        expect(scale(0, 100, { ...bubble, count: "range" }, 4.5)).toBe("-50..150/50");
        expect(scale(0, 100, bubble)).toBe("-20..120/20");
        expect(scale(20, 90, bubble)).toBe("0..120/20");
        expect(scale(95, 100, { ...bubble, count: "range" }, 4.5)).toBe("94..102/2");
        // Across 0 the room alone pads the data, 15% of its range a side (#620): -8..7 -12..10 by 2, -6..6 -8..8 by 2,
        // -80..70 -120..100 by 20 and -15..2 -20..5 by 5, where 5% more would round out a step further.
        expect(scale(-8, 7, bubble)).toBe("-12..10/2");
        expect(scale(-6, 6, bubble)).toBe("-8..8/2");
        expect(scale(-80, 70, bubble)).toBe("-120..100/20");
        expect(scale(-15, 2, bubble)).toBe("-20..5/5");
        // A minimum the room carries below 0 stays and is padded across it: x 0.4..3.6 is -0.08..4.08, not 0..4.14.
        expect(scale(0.4, 3.6, { ...bubble, count: "range" }, 3)).toBe("-2..6/2");
        expect(scale(1, 4, bubble, 5)).toBe("0..5/1");
    });

    test("takes a longer step where the room's steps run out, counted between the bounds or over the padded range", () => {
        // 3, 9, 16, 12 on a short column axis: 0..18 by 2 is 9 steps, 0..20 by 5 four.
        expect(scale(0, 16, column, 9)).toBe("0..18/2");
        expect(scale(0, 16, column, 8.9)).toBe("0..20/5");
        // -2..7 pads to 9.9 steps of 1, rounded out to 11.
        expect(scale(-2, 7, column, 11)).toBe("-3..8/1");
        expect(scale(-2, 7, { ...column, count: "range" }, 10)).toBe("-3..8/1");
        expect(scale(-2, 7, column, 10)).toBe("-4..8/2");
        // A bar3D axis pads by nothing: 3, 5, 4, 7 is 0..7 by 1.
        expect(scale(2, 7, { ...column, ratio: 0 }, 7)).toBe("0..7/1");
    });

    test("counts the steps from a fixed bound and keeps a fixed step", () => {
        expect(paddedAxisScale(3, 9, { kind: "minimum", value: 1 }, column, { kind: "room", steps: 100 })).toEqual({ minimum: 1, maximum: 10, interval: 1 });
        expect(paddedAxisScale(3, 9, { kind: "maximum", value: 20 }, column, { kind: "room", steps: 100 })).toEqual({ minimum: 0, maximum: 20, interval: 2 });
        expect(paddedAxisScale(3, 9, { kind: "none" }, column, { kind: "fixed", interval: 4 })).toEqual({ minimum: 0, maximum: 12, interval: 4 });
    });
});
