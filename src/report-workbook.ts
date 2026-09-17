import { Decimal } from "decimal.js";
import ExcelJS from "exceljs";

import type { AiUsageItem, CostItem, ReportMetadata } from "./types.js";

Decimal.set({ precision: 50, rounding: Decimal.ROUND_HALF_EVEN });

const COLORS = {
  indigo: "3730A3",
  blue: "3B618E",
  azure: "0078D4",
  amber: "D97706",
  teal: "0D9488",
  gray: "64748B",
  ink: "0F172A",
  border: "D8E0EA",
  page: "F8FAFC",
  panel: "FFFFFF",
  band: "F1F5F9",
  project: "E8EEF8",
  white: "FFFFFF",
};

const REPORT_HEADERS = [
  "SI No.",
  "Resource Name",
  "Resource Group",
  "Resource Type",
  "Billing Model",
  "Project",
  "Application",
  "Environment",
  "Subscription",
  "Location",
  "Cost (USD)",
];

interface Resource {
  resourceName: string;
  resourceGroup: string;
  resourceType: string;
  billingModel: string;
  project: string;
  application: string;
  environment: string;
  subscription: string;
  location: string;
  cost: Decimal;
}

function text(value: unknown, fallback = "Unmapped"): string {
  const result = String(value ?? "").trim();
  return result || fallback;
}

function decimal(value: unknown): Decimal {
  try {
    return new Decimal(String(value ?? 0));
  } catch {
    return new Decimal(0);
  }
}

function presented(value: Decimal): number {
  return value.toDecimalPlaces(2, Decimal.ROUND_HALF_EVEN).toNumber();
}

function resourcesFrom(items: CostItem[]): Resource[] {
  return items.map((item) => ({
    resourceName: text(item.resource_name, ""),
    resourceGroup: text(item.resource_group),
    resourceType: text(item.resource_type),
    billingModel: text(item.billing_model),
    project: text(item.resolved_project ?? item.project),
    application: text(item.resolved_application ?? item.application),
    environment: text(item.resolved_environment ?? item.environment),
    subscription: text(item.subscription),
    location: text(item.location, ""),
    cost: decimal(item.cost_usd),
  }));
}

function sum(values: Iterable<Decimal>): Decimal {
  let result = new Decimal(0);
  for (const value of values) result = result.plus(value);
  return result;
}

function addTo(map: Map<string, Decimal>, key: string, amount: Decimal): void {
  map.set(key, (map.get(key) ?? new Decimal(0)).plus(amount));
}

function ranked(map: Map<string, Decimal>): Array<[string, Decimal]> {
  return [...map.entries()].sort((left, right) => right[1].comparedTo(left[1]));
}

function percent(value: Decimal, total: Decimal): number {
  return total.isZero() ? 0 : value.dividedBy(total).toNumber();
}

function border(): Partial<ExcelJS.Borders> {
  const side: Partial<ExcelJS.Border> = { style: "thin", color: { argb: COLORS.border } };
  return { top: side, left: side, bottom: side, right: side };
}

function styleHeader(row: ExcelJS.Row, from: number, to: number, fill = COLORS.indigo): void {
  row.height = 22;
  for (let column = from; column <= to; column += 1) {
    const cell = row.getCell(column);
    cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: COLORS.white } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fill } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = border();
  }
}

function styleBody(
  sheet: ExcelJS.Worksheet,
  startRow: number,
  endRow: number,
  endColumn: number,
): void {
  for (let row = startRow; row <= endRow; row += 1) {
    for (let column = 1; column <= endColumn; column += 1) {
      const cell = sheet.getCell(row, column);
      cell.font = { name: "Calibri", size: 11, color: { argb: COLORS.ink } };
      cell.border = border();
      cell.alignment = { vertical: "middle" };
      if ((row - startRow) % 2 === 1) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.band } };
      }
    }
  }
}

function configureSheet(sheet: ExcelJS.Worksheet, freezeRows = 1): void {
  sheet.properties.defaultRowHeight = 18;
  sheet.views = [
    freezeRows
      ? { state: "frozen", ySplit: freezeRows, showGridLines: false }
      : { state: "normal", showGridLines: false },
  ];
  sheet.pageSetup = {
    orientation: "landscape",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.25, right: 0.25, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
  };
}

function setWidths(sheet: ExcelJS.Worksheet, widths: number[]): void {
  widths.forEach((width, index) => {
    sheet.getColumn(index + 1).width = width;
  });
}

function currencyColumns(sheet: ExcelJS.Worksheet, columns: number[], start: number, end: number): void {
  for (const column of columns) {
    for (let row = start; row <= end; row += 1) sheet.getCell(row, column).numFmt = "$#,##0.00";
  }
}

function percentageColumns(sheet: ExcelJS.Worksheet, columns: number[], start: number, end: number): void {
  for (const column of columns) {
    for (let row = start; row <= end; row += 1) sheet.getCell(row, column).numFmt = "0.00%";
  }
}

function formatPeriod(start: string, end: string): string {
  const startDate = new Date(`${start}T00:00:00Z`);
  const endDate = new Date(`${end}T00:00:00Z`);
  const startText = startDate.toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    timeZone: "UTC",
  });
  const endText = endDate.toLocaleString("en-US", {
    month: "short",
    day: "numeric",
    year: "numeric",
    timeZone: "UTC",
  });
  return `${startText} - ${endText}`;
}

function addOverview(
  workbook: ExcelJS.Workbook,
  resources: Resource[],
  metadata: ReportMetadata,
): void {
  const sheet = workbook.addWorksheet("Overview");
  configureSheet(sheet, 0);
  setWidths(sheet, [2.5, 26, 17, 9, 13, 9, 15, 9, 9, 9, 9, 9, 9, 2.5]);
  for (let row = 1; row <= 33; row += 1) {
    sheet.getRow(row).height = 18;
    for (let column = 1; column <= 14; column += 1) {
      sheet.getCell(row, column).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: COLORS.page },
      };
    }
  }
  sheet.getRow(1).height = 28;
  sheet.mergeCells("B1:M2");
  const title = sheet.getCell("B1");
  title.value = "Azure Cost Report";
  title.font = { name: "Calibri", size: 22, bold: true, color: { argb: COLORS.white } };
  title.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.blue } };
  title.alignment = { vertical: "middle", horizontal: "left", indent: 1 };
  for (let row = 1; row <= 2; row += 1) {
    for (let column = 2; column <= 13; column += 1) {
      sheet.getCell(row, column).fill = title.fill;
    }
  }
  sheet.mergeCells("B3:M3");
  sheet.getCell("B3").value = "Breakdown of Azure spend by project, application, and resource";
  sheet.getCell("B3").font = { name: "Calibri", size: 11, color: { argb: COLORS.gray } };

  const subscriptions = [...new Set(resources.map((item) => item.subscription))].sort();
  const subscription = subscriptions.length === 1 ? subscriptions[0] : "All subscriptions";
  const period = formatPeriod(metadata.period_start, metadata.period_end);
  const info: Array<[string, string, string, string]> = [
    ["B5", "Subscription:", "C5:E5", subscription],
    ["B6", "Currency:", "C6:E6", "USD"],
    ["B7", "Scope:", "C7:E7", "All resources"],
    ["G5", "Billing period:", "H5:M5", period],
    ["G6", "Generated:", "H6:M6", metadata.generated_at],
    ["G7", "Line items:", "H7:M7", String(resources.length)],
  ];
  for (const [labelCell, label, valueRange, value] of info) {
    sheet.getCell(labelCell).value = label;
    sheet.getCell(labelCell).font = { name: "Calibri", size: 10, bold: true, color: { argb: COLORS.gray } };
    sheet.mergeCells(valueRange);
    const valueCell = sheet.getCell(valueRange.split(":")[0]);
    valueCell.value = value;
    valueCell.font = { name: "Calibri", size: 11, bold: true, color: { argb: COLORS.ink } };
  }

  sheet.mergeCells("B9:M9");
  sheet.getCell("B9").value = "Key Metrics";
  sheet.getCell("B9").font = { name: "Calibri", size: 13, bold: true, color: { argb: COLORS.ink } };
  const total = sum(resources.map((item) => item.cost));
  const highest = [...resources].sort((left, right) => right.cost.comparedTo(left.cost))[0];
  sheet.mergeCells("B11:E11");
  sheet.mergeCells("B12:E13");
  sheet.mergeCells("B14:E14");
  sheet.mergeCells("F11:M11");
  sheet.mergeCells("F12:M13");
  sheet.mergeCells("F14:M14");
  sheet.getCell("B11").value = "TOTAL COST";
  sheet.getCell("F11").value = "HIGHEST COST ITEM";
  for (const address of ["B11", "F11"]) {
    const cell = sheet.getCell(address);
    cell.font = { name: "Calibri", size: 10, bold: true, color: { argb: COLORS.gray } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.panel } };
  }
  sheet.getCell("B12").value = presented(total);
  sheet.getCell("B12").numFmt = "$#,##0.00";
  sheet.getCell("B12").font = { name: "Calibri", size: 22, bold: true, color: { argb: COLORS.azure } };
  sheet.getCell("B14").value = `${period} | ${resources.length} line items`;
  sheet.getCell("F12").value = highest ? presented(highest.cost) : 0;
  sheet.getCell("F12").numFmt = "$#,##0.00";
  sheet.getCell("F12").font = { name: "Calibri", size: 22, bold: true, color: { argb: COLORS.amber } };
  sheet.getCell("F14").value = highest ? `${highest.application} - ${highest.resourceType}` : "No data";
  for (const address of ["B14", "F14"]) {
    sheet.getCell(address).font = { name: "Calibri", size: 10, color: { argb: COLORS.gray } };
  }

  sheet.mergeCells("B16:M16");
  sheet.getCell("B16").value = "Cost by Billing Model";
  sheet.getCell("B16").font = { name: "Calibri", size: 13, bold: true, color: { argb: COLORS.ink } };
  sheet.mergeCells("B17:M17");
  sheet.getCell("B17").value =
    "Fixed = provisioned capacity. Usage-Based = consumption-metered. Other = hybrid, one-off, or unclassified.";
  sheet.getCell("B17").font = { name: "Calibri", size: 10, color: { argb: COLORS.gray } };
  const buckets = new Map<string, Decimal>([
    ["Fixed", new Decimal(0)],
    ["Usage-Based", new Decimal(0)],
    ["Other", new Decimal(0)],
  ]);
  for (const resource of resources) {
    const bucket = resource.billingModel === "Fixed" || resource.billingModel === "Usage-Based"
      ? resource.billingModel
      : "Other";
    addTo(buckets, bucket, resource.cost);
  }
  const cards: Array<[string, string, string, string]> = [
    ["B18:E18", "B19:E22", "Fixed", COLORS.indigo],
    ["F18:I18", "F19:I22", "Usage-Based", COLORS.teal],
    ["J18:M18", "J19:M22", "Other", COLORS.gray],
  ];
  for (const [labelRange, valueRange, bucket, color] of cards) {
    sheet.mergeCells(labelRange);
    sheet.mergeCells(valueRange);
    const label = sheet.getCell(labelRange.split(":")[0]);
    label.value = bucket.toUpperCase();
    label.font = { name: "Calibri", size: 10, bold: true, color: { argb: COLORS.white } };
    label.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color } };
    label.alignment = { horizontal: "center", vertical: "middle" };
    const value = sheet.getCell(valueRange.split(":")[0]);
    value.value = presented(buckets.get(bucket) ?? new Decimal(0));
    value.numFmt = "$#,##0.00";
    value.font = { name: "Calibri", size: 18, bold: true, color: { argb: color } };
    value.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.panel } };
    value.alignment = { horizontal: "center", vertical: "middle" };
  }

  const unmapped = resources.filter(
    (item) => item.project === "Unmapped" || item.application === "Unmapped",
  );
  sheet.mergeCells("B27:M28");
  sheet.getCell("B27").value = "Report generated from canonical Supabase cost and AI usage records.";
  sheet.getCell("B27").font = { name: "Calibri", size: 10, color: { argb: COLORS.gray } };
  sheet.mergeCells("B29:M31");
  sheet.getCell("B29").value = unmapped.length
    ? `${unmapped.length} line items are not mapped to a configured project/application.`
    : "All line items are mapped to a project/application.";
  sheet.getCell("B29").font = { name: "Calibri", size: 10, italic: true, color: { argb: COLORS.gray } };
  sheet.getCell("B29").alignment = { vertical: "middle", wrapText: true };
}

function addApplicationSummary(workbook: ExcelJS.Workbook, resources: Resource[]): void {
  const sheet = workbook.addWorksheet("Cost per Application");
  configureSheet(sheet);
  sheet.addRow(["Project", "Application", "Total Cost (USD)", "% of Total Cost"]);
  styleHeader(sheet.getRow(1), 1, 4);
  const hierarchy = new Map<string, Map<string, Decimal>>();
  for (const item of resources) {
    if (!hierarchy.has(item.project)) hierarchy.set(item.project, new Map());
    addTo(hierarchy.get(item.project)!, item.application, item.cost);
  }
  const total = sum(resources.map((item) => item.cost));
  const projects = [...hierarchy.entries()]
    .map(([project, applications]) => ({ project, applications, total: sum(applications.values()) }))
    .sort((left, right) => right.total.comparedTo(left.total));
  for (const project of projects) {
    const projectRow = sheet.addRow([project.project, "", presented(project.total), percent(project.total, total)]);
    for (let column = 1; column <= 4; column += 1) {
      projectRow.getCell(column).font = { name: "Calibri", size: 11, bold: true, color: { argb: COLORS.ink } };
      projectRow.getCell(column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.project } };
      projectRow.getCell(column).border = border();
    }
    for (const [application, cost] of ranked(project.applications)) {
      const row = sheet.addRow(["", application, presented(cost), percent(cost, total)]);
      row.getCell(2).alignment = { indent: 1 };
    }
    sheet.addRow([]);
  }
  const grand = sheet.addRow(["Grand Total", "", presented(total), total.isZero() ? 0 : 1]);
  grand.font = { name: "Calibri", size: 11, bold: true, color: { argb: COLORS.white } };
  for (let column = 1; column <= 4; column += 1) {
    grand.getCell(column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.indigo } };
    grand.getCell(column).border = border();
  }
  styleBody(sheet, 2, sheet.rowCount - 1, 4);
  for (let row = 2; row < sheet.rowCount; row += 1) {
    if (sheet.getCell(row, 1).value && !sheet.getCell(row, 2).value) {
      for (let column = 1; column <= 4; column += 1) {
        const cell = sheet.getCell(row, column);
        cell.font = { name: "Calibri", size: 11, bold: true, color: { argb: COLORS.ink } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.project } };
      }
    }
  }
  currencyColumns(sheet, [3], 2, sheet.rowCount);
  percentageColumns(sheet, [4], 2, sheet.rowCount);
  setWidths(sheet, [28, 42, 18, 18]);
  sheet.autoFilter = `A1:D${sheet.rowCount}`;
}

function addResourceGroupSummary(workbook: ExcelJS.Workbook, resources: Resource[]): void {
  const sheet = workbook.addWorksheet("Cost per Resource Group");
  configureSheet(sheet);
  sheet.addRow([
    "Subscription",
    "Resource Group",
    "Cost (USD)",
    "% of Total Cost",
    "",
    "Environment",
    "Cost (USD)",
    "% of Total Cost",
  ]);
  styleHeader(sheet.getRow(1), 1, 4);
  styleHeader(sheet.getRow(1), 6, 8, COLORS.teal);
  const hierarchy = new Map<string, Map<string, Decimal>>();
  const environments = new Map<string, Decimal>();
  for (const item of resources) {
    if (!hierarchy.has(item.subscription)) hierarchy.set(item.subscription, new Map());
    addTo(hierarchy.get(item.subscription)!, item.resourceGroup, item.cost);
    addTo(environments, item.environment, item.cost);
  }
  const total = sum(resources.map((item) => item.cost));
  const subscriptions = [...hierarchy.entries()]
    .map(([subscription, groups]) => ({ subscription, groups, total: sum(groups.values()) }))
    .sort((left, right) => right.total.comparedTo(left.total));
  const leftRows: unknown[][] = [];
  for (const subscription of subscriptions) {
    leftRows.push([subscription.subscription, "", presented(subscription.total), percent(subscription.total, total)]);
    for (const [group, cost] of ranked(subscription.groups)) {
      leftRows.push(["", group, presented(cost), percent(cost, total)]);
    }
    leftRows.push([]);
  }
  leftRows.push(["Grand Total", "", presented(total), total.isZero() ? 0 : 1]);
  const environmentRows = ranked(environments).map(([environment, cost]) => [
    environment,
    presented(cost),
    percent(cost, total),
  ]);
  environmentRows.push(["Grand Total", presented(total), total.isZero() ? 0 : 1]);
  const rowCount = Math.max(leftRows.length, environmentRows.length);
  for (let index = 0; index < rowCount; index += 1) {
    sheet.addRow([...(leftRows[index] ?? []), "", ...(environmentRows[index] ?? [])]);
  }
  styleBody(sheet, 2, sheet.rowCount, 8);
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    if (sheet.getCell(row, 1).value && !sheet.getCell(row, 2).value) {
      for (let column = 1; column <= 4; column += 1) {
        sheet.getCell(row, column).font = { name: "Calibri", size: 11, bold: true };
        sheet.getCell(row, column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.project } };
      }
    }
  }
  currencyColumns(sheet, [3, 7], 2, sheet.rowCount);
  percentageColumns(sheet, [4, 8], 2, sheet.rowCount);
  setWidths(sheet, [32, 32, 17, 18, 4, 26, 17, 18]);
  sheet.autoFilter = `A1:D${sheet.rowCount}`;
}

function resourceRow(rank: number, resource: Resource): unknown[] {
  return [
    rank,
    resource.resourceName,
    resource.resourceGroup,
    resource.resourceType,
    resource.billingModel,
    resource.project,
    resource.application,
    resource.environment,
    resource.subscription,
    resource.location,
    presented(resource.cost),
  ];
}

function addTopResources(workbook: ExcelJS.Workbook, resources: Resource[]): void {
  const sheet = workbook.addWorksheet("Top 10 Resources");
  configureSheet(sheet, 3);
  sheet.mergeCells("A1:K1");
  sheet.getCell("A1").value = "Top 10 Resources by Cost";
  sheet.getCell("A1").font = { name: "Calibri", size: 15, bold: true, color: { argb: COLORS.indigo } };
  sheet.getRow(1).height = 26;
  sheet.addRow([]);
  sheet.addRow(REPORT_HEADERS.map((header) => (header === "SI No." ? "Rank" : header)));
  styleHeader(sheet.getRow(3), 1, 11);
  const top = [...resources].sort((left, right) => right.cost.comparedTo(left.cost)).slice(0, 10);
  top.forEach((resource, index) => sheet.addRow(resourceRow(index + 1, resource)));
  styleBody(sheet, 4, sheet.rowCount, 11);
  currencyColumns(sheet, [11], 4, sheet.rowCount);
  setWidths(sheet, [8, 34, 30, 42, 18, 26, 45, 20, 30, 22, 16]);
  sheet.autoFilter = `A3:K${sheet.rowCount}`;
}

function addResourceTypeSummary(workbook: ExcelJS.Workbook, resources: Resource[]): void {
  const sheet = workbook.addWorksheet("Cost per Resource Type");
  configureSheet(sheet);
  sheet.addRow(["Resource Type", "Cost (USD)", "", "Billing Model", "Cost (USD)", "% of Total Cost"]);
  styleHeader(sheet.getRow(1), 1, 2);
  styleHeader(sheet.getRow(1), 4, 6, COLORS.teal);
  const types = new Map<string, Decimal>();
  const billing = new Map<string, Decimal>();
  for (const item of resources) {
    addTo(types, item.resourceType, item.cost);
    addTo(billing, item.billingModel, item.cost);
  }
  const total = sum(resources.map((item) => item.cost));
  const typeRows: unknown[][] = ranked(types).map(([type, cost]) => [type, presented(cost)]);
  typeRows.push(["Grand Total", presented(total)]);
  const billingRows: unknown[][] = ranked(billing).map(([model, cost]) => [
    model,
    presented(cost),
    percent(cost, total),
  ]);
  billingRows.push(["Grand Total", presented(total), total.isZero() ? 0 : 1]);
  const rows = Math.max(typeRows.length, billingRows.length);
  for (let index = 0; index < rows; index += 1) {
    sheet.addRow([...(typeRows[index] ?? []), "", ...(billingRows[index] ?? [])]);
  }
  styleBody(sheet, 2, sheet.rowCount, 6);
  currencyColumns(sheet, [2, 5], 2, sheet.rowCount);
  percentageColumns(sheet, [6], 2, sheet.rowCount);
  setWidths(sheet, [42, 18, 4, 22, 18, 18]);
  for (const column of [1, 4]) {
    const row = sheet.rowCount;
    if (sheet.getCell(row, column).value === "Grand Total") sheet.getRow(row).font = { bold: true };
  }
}

function addCostReport(workbook: ExcelJS.Workbook, resources: Resource[]): void {
  const sheet = workbook.addWorksheet("Cost Report");
  configureSheet(sheet);
  sheet.addRow(REPORT_HEADERS);
  styleHeader(sheet.getRow(1), 1, 11);
  const sorted = [...resources].sort((left, right) =>
    left.project.localeCompare(right.project, undefined, { sensitivity: "base" }) ||
    left.application.localeCompare(right.application, undefined, { sensitivity: "base" }) ||
    right.cost.comparedTo(left.cost),
  );
  sorted.forEach((resource, index) => sheet.addRow(resourceRow(index + 1, resource)));
  styleBody(sheet, 2, sheet.rowCount, 11);
  let previousProject = "";
  let band = false;
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    const project = String(sheet.getCell(row, 6).value ?? "");
    if (project !== previousProject) {
      band = !band;
      previousProject = project;
    }
    if (band) {
      for (let column = 1; column <= 11; column += 1) {
        sheet.getCell(row, column).fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLORS.band } };
      }
    }
  }
  currencyColumns(sheet, [11], 2, sheet.rowCount);
  setWidths(sheet, [8, 34, 30, 34, 18, 24, 32, 20, 30, 20, 16]);
  sheet.autoFilter = `A1:K${sheet.rowCount}`;
}

function foundryModel(resourceType: string): string | undefined {
  return /^Foundry\s*\((.+)\)$/i.exec(resourceType.trim())?.[1].trim();
}

function addAiUsage(
  workbook: ExcelJS.Workbook,
  costItems: CostItem[],
  usageItems: AiUsageItem[],
): void {
  if (!usageItems.length) return;
  const sheet = workbook.addWorksheet("AI Token Usage");
  configureSheet(sheet);
  const headers = [
    "Subscription",
    "Foundry Resource Name",
    "Resource Group",
    "Application",
    "Model Deployment Name",
    "Model Type",
    "Model Version",
    "Input Tokens",
    "Output Tokens",
    "Total Tokens",
    "Requests",
    "Cost",
  ];
  sheet.addRow(headers);
  styleHeader(sheet.getRow(1), 1, 12);
  const applications = new Map<string, string>();
  const costs = new Map<string, Decimal>();
  for (const item of costItems) {
    const base = [item.subscription, item.resource_group, item.resource_name]
      .map((value) => String(value ?? "").trim().toLowerCase())
      .join("\u0000");
    applications.set(base, text(item.resolved_application ?? item.application, ""));
    const model = foundryModel(item.resource_type);
    if (model) costs.set(`${base}\u0000${model.toLowerCase()}`, decimal(item.cost_usd));
  }
  for (const item of usageItems.filter((row) => Number(row.total_tokens) !== 0)) {
    const base = [item.subscription, item.resource_group, item.foundry_resource_name]
      .map((value) => value.trim().toLowerCase())
      .join("\u0000");
    const cost = costs.get(`${base}\u0000${item.model_type.trim().toLowerCase()}`);
    sheet.addRow([
      item.subscription,
      item.foundry_resource_name,
      item.resource_group,
      applications.get(base) || null,
      item.model_deployment_name,
      item.model_type,
      item.model_version,
      Number(item.input_tokens),
      Number(item.output_tokens),
      Number(item.total_tokens),
      Number(item.requests),
      cost ? presented(cost) : null,
    ]);
  }
  styleBody(sheet, 2, sheet.rowCount, 12);
  for (let row = 2; row <= sheet.rowCount; row += 1) {
    for (let column = 8; column <= 11; column += 1) sheet.getCell(row, column).numFmt = "#,##0";
  }
  currencyColumns(sheet, [12], 2, sheet.rowCount);
  setWidths(sheet, [36, 32, 28, 34, 38, 24, 18, 18, 18, 18, 14, 16]);
  sheet.autoFilter = `A1:L${sheet.rowCount}`;
}

export async function buildReportWorkbook(
  metadata: ReportMetadata,
  costItems: CostItem[],
  aiUsage: AiUsageItem[],
): Promise<Buffer> {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Cirrus Cost Reporting";
  workbook.created = new Date();
  workbook.modified = new Date();
  workbook.calcProperties.fullCalcOnLoad = true;
  const resources = resourcesFrom(costItems);
  addOverview(workbook, resources, metadata);
  addApplicationSummary(workbook, resources);
  addResourceGroupSummary(workbook, resources);
  addTopResources(workbook, resources);
  addResourceTypeSummary(workbook, resources);
  addCostReport(workbook, resources);
  addAiUsage(workbook, costItems, aiUsage);
  return Buffer.from(await workbook.xlsx.writeBuffer());
}
