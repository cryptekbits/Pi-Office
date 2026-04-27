import type { OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import {
  supportsRequirementSet,
  isRecord,
  trimString,
  toNumber,
  toBoolean,
  splitSheetAddress,
  withSheetName,
} from "../shared";
import {
  resolveExcelWorksheet,
  resolveExcelRange,
  resolveExcelChart,
} from "../excel-targets";

function applyExcelChartAxisOptions(axis: Excel.ChartAxis, value: unknown): void {
  if (!isRecord(value)) {
    return;
  }

  const title = trimString(value.title) ?? trimString(value.text);
  const titleVisible = toBoolean(value.titleVisible);
  const visible = toBoolean(value.visible);
  const displayUnit = trimString(value.displayUnit);
  const numberFormat = trimString(value.numberFormat);
  const minimum = toNumber(value.minimum);
  const maximum = toNumber(value.maximum);
  const majorUnit = toNumber(value.majorUnit);
  const minorUnit = toNumber(value.minorUnit);
  const majorGridlinesVisible = toBoolean(value.majorGridlinesVisible);
  const minorGridlinesVisible = toBoolean(value.minorGridlinesVisible);
  const reversePlotOrder = toBoolean(value.reversePlotOrder);
  const logBase = toNumber(value.logBase);

  if (typeof visible === "boolean") axis.visible = visible;
  if (title) {
    axis.title.text = title;
    axis.title.visible = true;
  } else if (typeof titleVisible === "boolean") {
    axis.title.visible = titleVisible;
  }
  if (displayUnit) axis.displayUnit = displayUnit as Excel.ChartAxisDisplayUnit;
  if (numberFormat) {
    axis.numberFormat = numberFormat;
    if (supportsRequirementSet("ExcelApi", "1.9")) {
      axis.linkNumberFormat = false;
    }
  }
  if (typeof minimum === "number") axis.minimum = minimum;
  if (typeof maximum === "number") axis.maximum = maximum;
  if (typeof majorUnit === "number") axis.majorUnit = majorUnit;
  if (typeof minorUnit === "number") axis.minorUnit = minorUnit;
  if (typeof majorGridlinesVisible === "boolean") axis.majorGridlines.visible = majorGridlinesVisible;
  if (typeof minorGridlinesVisible === "boolean") axis.minorGridlines.visible = minorGridlinesVisible;
  if (typeof reversePlotOrder === "boolean") axis.reversePlotOrder = reversePlotOrder;
  if (typeof logBase === "number") axis.logBase = logBase;
}

function applyExcelChartDataLabelOptions(labels: Excel.ChartDataLabels, value: unknown): void {
  if (!isRecord(value)) {
    return;
  }

  const position = trimString(value.position);
  const separator = trimString(value.separator);
  const showValue = toBoolean(value.showValue);
  const showCategoryName = toBoolean(value.showCategoryName);
  const showSeriesName = toBoolean(value.showSeriesName);
  const showLegendKey = toBoolean(value.showLegendKey);
  const showPercentage = toBoolean(value.showPercentage);
  const showBubbleSize = toBoolean(value.showBubbleSize);
  const showLeaderLines = toBoolean(value.showLeaderLines);
  const numberFormat = trimString(value.numberFormat);
  const fontColor = trimString(value.fontColor);
  const fontSize = toNumber(value.fontSize);
  const bold = toBoolean(value.bold);
  const italic = toBoolean(value.italic);
  const fillColor = trimString(value.fillColor);
  const borderColor = trimString(value.borderColor);
  const borderWeight = toNumber(value.borderWeight);

  if (position) labels.position = position as Excel.ChartDataLabelPosition;
  if (separator) labels.separator = separator;
  if (typeof showValue === "boolean") labels.showValue = showValue;
  if (typeof showCategoryName === "boolean") labels.showCategoryName = showCategoryName;
  if (typeof showSeriesName === "boolean") labels.showSeriesName = showSeriesName;
  if (typeof showLegendKey === "boolean") labels.showLegendKey = showLegendKey;
  if (typeof showPercentage === "boolean") labels.showPercentage = showPercentage;
  if (typeof showBubbleSize === "boolean") labels.showBubbleSize = showBubbleSize;
  if (typeof showLeaderLines === "boolean" && supportsRequirementSet("ExcelApi", "1.19")) {
    labels.showLeaderLines = showLeaderLines;
  }
  if (numberFormat) {
    labels.numberFormat = numberFormat;
    if (supportsRequirementSet("ExcelApi", "1.9")) {
      labels.linkNumberFormat = false;
    }
  }
  if (fontColor) labels.format.font.color = fontColor;
  if (typeof fontSize === "number") labels.format.font.size = fontSize;
  if (typeof bold === "boolean") labels.format.font.bold = bold;
  if (typeof italic === "boolean") labels.format.font.italic = italic;
  if (fillColor) labels.format.fill.setSolidColor(fillColor);
  if (borderColor) labels.format.border.color = borderColor;
  if (typeof borderWeight === "number") labels.format.border.weight = borderWeight;
}

async function applyExcelChartConfiguration(
  context: Excel.RequestContext,
  chart: Excel.Chart,
  value: Record<string, unknown>,
): Promise<void> {
  const source = trimString(value.source);
  if (source) {
    const parsed = splitSheetAddress(source, trimString(value.sheetName));
    const sourceWorksheet = resolveExcelWorksheet(context, { kind: "sheet", sheetName: parsed.sheetName }, true);
    chart.setData(sourceWorksheet.getRange(parsed.address ?? source), (trimString(value.seriesBy) ?? "Auto") as Excel.ChartSeriesBy);
  }

  const chartType = trimString(value.chartType);
  const title = trimString(value.title);
  const legendVisible = toBoolean(value.legendVisible);
  const legendPosition = trimString(value.legendPosition);
  const chartLeft = toNumber(value.left);
  const chartTop = toNumber(value.top);
  const chartWidth = toNumber(value.width);
  const chartHeight = toNumber(value.height);
  if (chartType) chart.chartType = chartType as Excel.ChartType;
  if (title) {
    chart.title.text = title;
    chart.title.visible = true;
  }
  if (typeof legendVisible === "boolean") chart.legend.visible = legendVisible;
  if (legendPosition) chart.legend.position = legendPosition as Excel.ChartLegendPosition;
  if (typeof chartLeft === "number") chart.left = chartLeft;
  if (typeof chartTop === "number") chart.top = chartTop;
  if (typeof chartWidth === "number") chart.width = chartWidth;
  if (typeof chartHeight === "number") chart.height = chartHeight;

  const categoryAxis: Record<string, unknown> = {};
  if (isRecord(value.categoryAxis)) Object.assign(categoryAxis, value.categoryAxis);
  if (value.categoryAxisTitle != null) categoryAxis.title = value.categoryAxisTitle;
  if (value.categoryAxisVisible != null) categoryAxis.visible = value.categoryAxisVisible;
  if (value.categoryAxisNumberFormat != null) categoryAxis.numberFormat = value.categoryAxisNumberFormat;
  if (Object.keys(categoryAxis).length) {
    applyExcelChartAxisOptions(chart.axes.categoryAxis, categoryAxis);
  }

  const valueAxis: Record<string, unknown> = {};
  if (isRecord(value.valueAxis)) Object.assign(valueAxis, value.valueAxis);
  if (value.valueAxisTitle != null) valueAxis.title = value.valueAxisTitle;
  if (value.valueAxisVisible != null) valueAxis.visible = value.valueAxisVisible;
  if (value.valueAxisDisplayUnit != null) valueAxis.displayUnit = value.valueAxisDisplayUnit;
  if (value.valueAxisMinimum != null) valueAxis.minimum = value.valueAxisMinimum;
  if (value.valueAxisMaximum != null) valueAxis.maximum = value.valueAxisMaximum;
  if (value.valueAxisMajorUnit != null) valueAxis.majorUnit = value.valueAxisMajorUnit;
  if (value.valueAxisMinorUnit != null) valueAxis.minorUnit = value.valueAxisMinorUnit;
  if (value.valueAxisMajorGridlinesVisible != null) valueAxis.majorGridlinesVisible = value.valueAxisMajorGridlinesVisible;
  if (value.valueAxisNumberFormat != null) valueAxis.numberFormat = value.valueAxisNumberFormat;
  if (Object.keys(valueAxis).length) {
    applyExcelChartAxisOptions(chart.axes.valueAxis, valueAxis);
  }

  const dataLabels: Record<string, unknown> = {};
  if (isRecord(value.dataLabels)) Object.assign(dataLabels, value.dataLabels);
  if (value.showDataLabels != null) dataLabels.visible = value.showDataLabels;
  if (value.dataLabelPosition != null) dataLabels.position = value.dataLabelPosition;
  if (value.dataLabelSeparator != null) dataLabels.separator = value.dataLabelSeparator;
  if (value.showDataLabelValue != null) dataLabels.showValue = value.showDataLabelValue;
  if (value.showDataLabelCategoryName != null) dataLabels.showCategoryName = value.showDataLabelCategoryName;
  if (value.showDataLabelSeriesName != null) dataLabels.showSeriesName = value.showDataLabelSeriesName;
  if (value.showDataLabelLegendKey != null) dataLabels.showLegendKey = value.showDataLabelLegendKey;
  if (value.showDataLabelPercentage != null) dataLabels.showPercentage = value.showDataLabelPercentage;
  if (value.showDataLabelBubbleSize != null) dataLabels.showBubbleSize = value.showDataLabelBubbleSize;
  if (value.showDataLabelLeaderLines != null) dataLabels.showLeaderLines = value.showDataLabelLeaderLines;
  if (value.dataLabelNumberFormat != null) dataLabels.numberFormat = value.dataLabelNumberFormat;
  if (value.dataLabelFontColor != null) dataLabels.fontColor = value.dataLabelFontColor;
  if (value.dataLabelFontSize != null) dataLabels.fontSize = value.dataLabelFontSize;
  if (Object.keys(dataLabels).length) {
    chart.series.load("items/name");
    await context.sync();
    const visible = toBoolean(dataLabels.visible);
    const shouldEnable = typeof visible === "boolean" ? visible : true;
    for (const series of chart.series.items) {
      series.hasDataLabels = shouldEnable;
      if (dataLabels.showLeaderLines != null && supportsRequirementSet("ExcelApi", "1.9")) {
        series.showLeaderLines = toBoolean(dataLabels.showLeaderLines) ?? false;
      }
    }
    applyExcelChartDataLabelOptions(chart.dataLabels, dataLabels);
  }
}

function escapeXmlText(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function escapeXmlAttribute(value: string): string {
  return escapeXmlText(value).replaceAll("\"", "&quot;").replaceAll("'", "&apos;");
}

export async function applyExcelCreateChartAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
  const source = trimString(options.source) ?? withSheetName(action.target?.sheetName, action.target?.address);
  const parsedSource = splitSheetAddress(source, action.target?.sheetName);
  const sourceRange = parsedSource.address
    ? resolveExcelWorksheet(context, { kind: "sheet", sheetName: parsedSource.sheetName }, true).getRange(parsedSource.address)
    : resolveExcelRange(context, action.target, true);
  const worksheet = resolveExcelWorksheet(context, action.target, true);
  const chart = worksheet.charts.add(
    (trimString(options.chartType) ?? "ColumnClustered") as Excel.ChartType,
    sourceRange,
    (trimString(options.seriesBy) ?? "Auto") as Excel.ChartSeriesBy,
  );
  const chartName = trimString(options.chartName);
  if (chartName) chart.name = chartName;
  const chartTitle = trimString(options.title);
  if (chartTitle) {
    chart.title.text = chartTitle;
    chart.title.visible = true;
  }
  const chartLeft = toNumber(options.left);
  const chartTop = toNumber(options.top);
  const chartWidth = toNumber(options.width);
  const chartHeight = toNumber(options.height);
  const legendVisible = toBoolean(options.legendVisible);
  if (typeof chartLeft === "number") chart.left = chartLeft;
  if (typeof chartTop === "number") chart.top = chartTop;
  if (typeof chartWidth === "number") chart.width = chartWidth;
  if (typeof chartHeight === "number") chart.height = chartHeight;
  if (typeof legendVisible === "boolean") chart.legend.visible = legendVisible;
  await applyExcelChartConfiguration(context, chart, options);
  chart.load("name,id");
  await context.sync();
  return { ok: true, host: "excel", action: type, chartName: chart.name, chartId: chart.id };
}

export async function applyExcelUpdateChartAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
  const chart = resolveExcelChart(context, action.target ?? { kind: "chart", chartName: trimString(options.chartName) ?? "" });
  await applyExcelChartConfiguration(context, chart, options);
  chart.load("name,id");
  await context.sync();
  return { ok: true, host: "excel", action: type, chartName: chart.name, chartId: chart.id };
}

export async function applyExcelExtractChartXmlAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
  options: Record<string, unknown>,
): Promise<unknown> {
  const worksheet = resolveExcelWorksheet(context, action.target, true);
  const charts = worksheet.charts;
  charts.load("items/name,items/id");
  worksheet.load("name");
  await context.sync();

  const targetChartName = trimString(action.target?.chartName) ?? trimString(options.chartName) ?? trimString(options.name);
  const targetChartId = trimString(options.chartId) ?? trimString(options.id);
  const targetChartIndex = toNumber(options.chartIndex);

  let chart: Excel.Chart | undefined;
  if (targetChartName) {
    chart = charts.getItem(targetChartName);
  } else if (targetChartId) {
    const matched = charts.items.find((candidate) => candidate.id === targetChartId);
    chart = matched ? charts.getItem(matched.name) : undefined;
  } else if (typeof targetChartIndex === "number" && Number.isInteger(targetChartIndex) && targetChartIndex > 0) {
    chart = charts.getItemAt(targetChartIndex - 1);
  } else if (charts.items.length > 0) {
    chart = charts.getItemAt(0);
  }

  if (!chart) {
    throw new Error("Excel chart XML extraction requires an existing chart target.");
  }

  chart.load("name,id,chartType,left,top,width,height");
  chart.title.load("text,visible");
  chart.legend.load("visible,position");
  chart.axes.categoryAxis.load("visible");
  chart.axes.valueAxis.load("visible");
  chart.series.load("items/name");
  await context.sync();

  const titleText = chart.title.visible ? trimString(chart.title.text) ?? "" : "";
  const seriesXml = chart.series.items
    .map((series, index) => {
      const name = trimString(series.name) ?? `Series ${index + 1}`;
      return `<series index="${index + 1}" name="${escapeXmlAttribute(name)}" />`;
    })
    .join("");

  const chartXml = [
    `<chart name="${escapeXmlAttribute(chart.name)}" id="${escapeXmlAttribute(chart.id)}" sheetName="${escapeXmlAttribute(worksheet.name)}" chartType="${escapeXmlAttribute(String(chart.chartType ?? ""))}">`,
    `<title visible="${chart.title.visible ? "true" : "false"}">${escapeXmlText(titleText)}</title>`,
    `<legend visible="${chart.legend.visible ? "true" : "false"}" position="${escapeXmlAttribute(String(chart.legend.position ?? ""))}" />`,
    "<axes>",
    `<category visible="${chart.axes.categoryAxis.visible ? "true" : "false"}" />`,
    `<value visible="${chart.axes.valueAxis.visible ? "true" : "false"}" />`,
    "</axes>",
    `<position left="${chart.left}" top="${chart.top}" width="${chart.width}" height="${chart.height}" />`,
    seriesXml ? `<seriesList>${seriesXml}</seriesList>` : "",
    "</chart>",
  ].join("");

  return {
    ok: true,
    host: "excel",
    action: type,
    sheetName: worksheet.name,
    chartName: chart.name,
    chartId: chart.id,
    chartType: chart.chartType,
    chartXml,
  };
}
