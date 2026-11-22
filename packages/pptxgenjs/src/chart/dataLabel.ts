import { PPTX_DEFAULTS } from "#src/defaults";
import type { NormalizedChartDataLabels } from "@pdflapas/model";
import { formatChartNumber } from "./format";

/** Whether a point of a chart other than pie or doughnut gets a label. */
export function showsPointLabel(labels: NormalizedChartDataLabels): boolean {
    return labels.showValue || labels.showCategory || labels.showSeriesName === true;
}

/**
 * The label of a point of a chart other than pie or doughnut: a scatter point's own label (`customLabel`) alone, as
 * PptxGenJS writes it as the point's whole `<c:tx>`, or the series name and the value joined as LibreOffice joins them.
 */
export function pointLabelText(labels: NormalizedChartDataLabels, seriesName: string, value: number, customLabel: { kind: "label"; text: string } | { kind: "none" }): string {
    if (labels.showCategory && customLabel.kind === "label") {
        return customLabel.text;
    }

    const fields: string[] = [];

    if (labels.showSeriesName) {
        fields.push(seriesName);
    }

    if (labels.showValue) {
        fields.push(formatChartNumber(value, labels.valueFormat));
    }

    return fields.join(labels.fieldSeparator ?? PPTX_DEFAULTS.chart.dataLabelFieldSeparator);
}
