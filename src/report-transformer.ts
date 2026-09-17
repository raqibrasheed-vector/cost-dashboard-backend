import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";

import { parse as parseCsv } from "csv-parse/sync";
import { Decimal } from "decimal.js";
import ExcelJS from "exceljs";

import {
  applicationOwnership,
  resourceApplications,
  resourceGroupApplications,
  resourceTypeBillingModels,
} from "./report-mappings.js";
import type { AiUsageItem, CostItem, ReportMetadata, TransformResult } from "./types.js";

Decimal.set({ precision: 50 });

export const AI_USAGE_COLUMNS = [
  "Subscription",
  "Foundry Resource Name",
  "Resource Group",
  "Model Deployment Name",
  "Model Type",
  "Model Version",
  "Input Tokens",
  "Output Tokens",
  "Total Tokens",
  "Requests",
] as const;

type SourceRow = Record<string, unknown>;

interface Resource {
  resource_name: string;
  resource_id: string;
  resource_group: string;
  resource_group_id: string;
  service: string;
  billing_model: string;
  project: string;
  application: string;
  environment: string;
  subscription: string;
  subscription_id: string;
  location: string;
  cost: Decimal;
}

interface Metadata {
  period_start: Date;
  period_end: Date;
  generated_at: Date;
}

function text(value: unknown): string {
  return String(value ?? "").trim();
}

function normalized(value: unknown): string {
  return text(value).toLocaleLowerCase("en-US");
}

function scalar(value: ExcelJS.CellValue): unknown {
  if (value && typeof value === "object" && !(value instanceof Date)) {
    if ("result" in value) return value.result;
    if ("richText" in value) return value.richText.map((part) => part.text).join("");
    if ("text" in value) return value.text;
  }
  return value;
}

function money(value: unknown): Decimal {
  const cleaned = text(value).replaceAll("$", "").replaceAll(",", "");
  if (!cleaned) return new Decimal(0);
  try {
    return new Decimal(cleaned);
  } catch {
    return new Decimal(0);
  }
}

function count(value: unknown, field: string): number {
  const cleaned = text(value || "0").replaceAll(",", "");
  let parsed: Decimal;
  try {
    parsed = new Decimal(cleaned || "0");
  } catch {
    throw new Error(`Invalid ${field}: ${String(value)}`);
  }
  if (parsed.isNegative() || !parsed.isInteger() || parsed.greaterThan(Number.MAX_SAFE_INTEGER)) {
    throw new Error(`Invalid ${field}: ${String(value)}`);
  }
  return parsed.toNumber();
}

function utcDate(year: number, month: number, day: number): Date {
  return new Date(Date.UTC(year, month - 1, day));
}

function parseDate(value: unknown, field: string): Date {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return utcDate(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate());
  }
  const raw = text(value);
  let match = /^(\d{4})-(\d{2})-(\d{2})(?:[ T].*)?$/.exec(raw);
  if (match) return utcDate(Number(match[1]), Number(match[2]), Number(match[3]));
  match = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(raw);
  if (match) return utcDate(Number(match[3]), Number(match[1]), Number(match[2]));
  match = /^([A-Za-z]+)\s+(\d{1,2}),\s*(\d{4})$/.exec(raw);
  if (match) {
    const month = new Date(`${match[1]} 1, 2000 UTC`).getUTCMonth() + 1;
    if (month > 0) return utcDate(Number(match[3]), month, Number(match[2]));
  }
  throw new Error(`Invalid date for ${field}: ${raw}`);
}

function isoDate(value: Date): string {
  return value.toISOString().slice(0, 10);
}

function monthStart(value: Date): string {
  return `${value.getUTCFullYear()}-${String(value.getUTCMonth() + 1).padStart(2, "0")}-01`;
}

function summaryMetadata(sheet: ExcelJS.Worksheet): Metadata {
  const labels: Record<string, keyof Metadata> = {
    "start date": "period_start",
    "end date": "period_end",
    generated: "generated_at",
  };
  const metadata: Partial<Metadata> = {};
  sheet.eachRow((row) => {
    const values = (row.values as ExcelJS.CellValue[]).slice(1).map(scalar);
    values.forEach((value, index) => {
      const raw = text(value);
      const label = raw.replace(/:$/, "").trim().toLowerCase();
      let key = labels[label];
      let adjacent = values[index + 1];
      if (!key) {
        const prefix = Object.keys(labels).find((item) => label.startsWith(`${item}:`));
        if (prefix) {
          key = labels[prefix];
          adjacent = raw.split(":", 2)[1];
        }
      }
      if (key && adjacent !== undefined && adjacent !== "") {
        metadata[key] = parseDate(adjacent, label);
      }
    });
  });
  for (const key of Object.values(labels)) {
    if (!metadata[key]) throw new Error(`Summary sheet is missing metadata for: ${key}`);
  }
  return metadata as Metadata;
}

function billingPeriod(value: string): [Date, Date] {
  const cleaned = value.replace(/[\u2013\u2014]/g, "-").replace(/\s+/g, " ").trim();
  const match = /^([A-Za-z]{3,9})\s+(\d{1,2})\s*-\s*([A-Za-z]{3,9})?\s*(\d{1,2}),\s*(\d{4})/.exec(cleaned);
  if (!match) throw new Error(`Invalid billing period: ${value}`);
  return [
    parseDate(`${match[1]} ${match[2]}, ${match[5]}`, "billing period start"),
    parseDate(`${match[3] || match[1]} ${match[4]}, ${match[5]}`, "billing period end"),
  ];
}

function generatedMetadata(workbook: ExcelJS.Workbook): Metadata {
  const overview = workbook.worksheets.find((sheet) => sheet.name.toLowerCase() === "overview");
  if (!overview) throw new Error("Generated reports require an Overview sheet.");
  let period: [Date, Date] | undefined;
  overview.eachRow((row) => {
    const values = (row.values as ExcelJS.CellValue[]).slice(1).map((value) => text(scalar(value)));
    values.forEach((value, index) => {
      if (value.replace(/:$/, "").toLowerCase() === "billing period" && values[index + 1]) {
        period = billingPeriod(values[index + 1]);
      }
    });
  });
  if (!period) throw new Error("Overview sheet is missing Billing period metadata.");
  const now = new Date();
  return {
    period_start: period[0],
    period_end: period[1],
    generated_at: utcDate(now.getUTCFullYear(), now.getUTCMonth() + 1, now.getUTCDate()),
  };
}

function foundryModel(source: string): string {
  const patterns: Array<[RegExp, string]> = [
    [/text[\s-]?embedding[\s-]?3[\s-]?large/i, "text-embedding-3-large"],
    [/text[\s-]?embedding[\s-]?3[\s-]?small/i, "text-embedding-3-small"],
    [/text[\s-]?embedding[\s-]?ada[\s-]?002/i, "text-embedding-ada-002"],
    [/embedding[\s-]?ada[\s-]?glbl/i, "text-embedding-ada-002"],
    [/embedding[\s-]?ada[\s-]?regional/i, "text-embedding-ada-002"],
    [/chat[\s-]?latest/i, "gpt-chat-latest"],
    [/5\.4/i, "gpt-5.4"],
    [/image\s*analysis/i, "Image Analysis"],
    [/foundry\s*tools?/i, "Foundry tools"],
    [/dall[\s-]?e[\s-]?3/i, "dall-e-3"],
    [/dall[\s-]?e[\s-]?2/i, "dall-e-2"],
    [/whisper/i, "whisper"],
    [/text[\s-]?moderation/i, "text-moderation"],
  ];
  for (const [pattern, label] of patterns) if (pattern.test(source)) return label;
  if (/gpt/i.test(source)) {
    const decimal = /\b(\d+\.\d+)\b/.exec(source);
    if (decimal) return `gpt-${decimal[1]}`;
    const matches = [...source.matchAll(/gpt[\s-]*(\d+)[\s-]*(o|mini|nano|turbo)?\b/gi)];
    const selected = [...matches].reverse().find((item) => item[2]) ?? matches.at(-1);
    if (selected) {
      const suffix = (selected[2] ?? "").toLowerCase();
      if (suffix === "o") return `gpt-${selected[1]}o`;
      return suffix ? `gpt-${selected[1]}-${suffix}` : `gpt-${selected[1]}`;
    }
  }
  const reasoning = /\bo(1|3|4)[\s-]?(mini|preview)?\b/i.exec(source);
  return reasoning ? `o${reasoning[1]}${reasoning[2] ? `-${reasoning[2].toLowerCase()}` : ""}` : "Unknown Model";
}

function resourceType(row: SourceRow): string {
  const type = text(row.ResourceType);
  if (type !== "Cognitive Services account") return type;
  if (text(row.ServiceName) !== "Foundry Models") return text(row.ServiceTier) || type;
  return `Foundry (${foundryModel(text(row.Product) || text(row.Meter))})`;
}

function billingModel(type: string): string {
  return resourceTypeBillingModels[type] ?? (type.toLowerCase().startsWith("foundry") ? "Usage-Based" : "Unmapped");
}

function resolveOwnership(resourceName: string, group: string) {
  const application = resourceApplications[resourceName] ?? resourceGroupApplications[group] ?? "Unmapped";
  const owner = applicationOwnership[application];
  return {
    application,
    project: owner?.project ?? "Unmapped",
    environment: owner?.environment ?? "Unmapped",
  };
}

function aggregateSourceRows(rows: SourceRow[]): Resource[] {
  const resources = new Map<string, Resource>();
  for (const row of rows) {
    const resourceName = text(row.Resource);
    if (!resourceName) continue;
    const group = text(row.ResourceGroupName);
    const subscription = text(row.SubscriptionName);
    const subscriptionId = text(row.SubscriptionId);
    const groupId = text(row.ResourceGroupId);
    const resourceId = text(row.ResourceId);
    const type = resourceType(row);
    const key = [subscriptionId || subscription, groupId || group, resourceId || resourceName, type].join("\u0000");
    let resource = resources.get(key);
    if (!resource) {
      const owner = resolveOwnership(resourceName, group);
      resource = {
        resource_name: resourceName,
        resource_id: resourceId,
        resource_group: group,
        resource_group_id: groupId,
        service: type,
        billing_model: billingModel(type),
        ...owner,
        subscription,
        subscription_id: subscriptionId,
        location: text(row.ResourceLocation),
        cost: new Decimal(0),
      };
      resources.set(key, resource);
    }
    resource.cost = resource.cost.plus(money(row.CostUSD || row.Cost));
  }
  return [...resources.values()].sort((left, right) => right.cost.comparedTo(left.cost));
}

function sheetRows(sheet: ExcelJS.Worksheet): SourceRow[] {
  const headerRow = sheet.getRow(1);
  const headers = (headerRow.values as ExcelJS.CellValue[]).slice(1).map((value) => text(scalar(value)));
  if (!headers.some(Boolean)) throw new Error(`The ${sheet.name} sheet is empty.`);
  const rows: SourceRow[] = [];
  for (let index = 2; index <= sheet.rowCount; index += 1) {
    const values = (sheet.getRow(index).values as ExcelJS.CellValue[]).slice(1).map(scalar);
    if (!values.some((value) => value !== null && value !== undefined && value !== "")) continue;
    rows.push(Object.fromEntries(headers.map((header, column) => [header, values[column]])));
  }
  return rows;
}

async function readCostWorkbook(filePath: string): Promise<[Resource[], Metadata]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const data = workbook.worksheets.find((sheet) => sheet.name.toLowerCase() === "data");
  const summary = workbook.worksheets.find((sheet) => sheet.name.toLowerCase() === "summary");
  if (data && summary) return [aggregateSourceRows(sheetRows(data)), summaryMetadata(summary)];

  const report = workbook.worksheets.find((sheet) => sheet.name.toLowerCase() === "cost report");
  if (!report) throw new Error(`${path.basename(filePath)}: expected Data/Summary or Cost Report sheets.`);
  const resources = sheetRows(report).map((row) => ({
    resource_name: text(row["Resource Name"]),
    resource_id: "",
    resource_group: text(row["Resource Group"]),
    resource_group_id: "",
    service: text(row["Resource Type"]),
    billing_model: text(row["Billing Model"]),
    project: text(row.Project),
    application: text(row.Application),
    environment: text(row.Environment),
    subscription: text(row.Subscription),
    subscription_id: "",
    location: text(row.Location),
    cost: money(row["Cost (USD)"]),
  }));
  return [resources, generatedMetadata(workbook)];
}

function combineResources(resources: Resource[]): Resource[] {
  const combined = new Map<string, Resource>();
  for (const resource of resources) {
    const key = [
      resource.subscription_id || resource.subscription,
      resource.resource_group_id || resource.resource_group,
      resource.resource_id || resource.resource_name,
      resource.service,
    ].join("\u0000");
    const current = combined.get(key);
    if (current) current.cost = current.cost.plus(resource.cost);
    else combined.set(key, { ...resource, cost: new Decimal(resource.cost) });
  }
  return [...combined.values()].sort((left, right) => right.cost.comparedTo(left.cost));
}

function mergeMetadata(items: Metadata[]): Metadata {
  const months = new Set(items.map((item) => `${item.period_start.getUTCFullYear()}-${item.period_start.getUTCMonth()}`));
  if (months.size !== 1) throw new Error("All cost workbooks must cover the same report month.");
  return {
    period_start: new Date(Math.min(...items.map((item) => item.period_start.getTime()))),
    period_end: new Date(Math.max(...items.map((item) => item.period_end.getTime()))),
    generated_at: new Date(Math.max(...items.map((item) => item.generated_at.getTime()))),
  };
}

function resourceKey(resource: Resource): string {
  return normalized(resource.resource_id) || [
    normalized(resource.subscription_id || resource.subscription),
    normalized(resource.resource_group_id || resource.resource_group),
    normalized(resource.resource_name),
  ].join("|");
}

function foundryType(type: string): string | undefined {
  return /^Foundry\s*\((.+)\)$/i.exec(type.trim())?.[1].trim();
}

function costPayload(resources: Resource[], metadata: Metadata): CostItem[] {
  return resources.map((resource) => ({
    report_month: monthStart(metadata.period_start),
    billing_period_start: isoDate(metadata.period_start),
    billing_period_end: isoDate(metadata.period_end),
    currency: "USD",
    resource_key: resourceKey(resource),
    resource_id: resource.resource_id || null,
    resource_name: resource.resource_name || null,
    resource_group: resource.resource_group || "Unmapped",
    resource_group_id: resource.resource_group_id || null,
    resource_type: resource.service || "Unmapped",
    billing_model: resource.billing_model || "Unmapped",
    project: resource.project || "Unmapped",
    application: resource.application || "Unmapped",
    environment: resource.environment || "Unmapped",
    subscription: resource.subscription || "Unmapped",
    subscription_id: resource.subscription_id || null,
    location: resource.location || null,
    cost_usd: resource.cost.toString(),
  }));
}

function aiPayload(rows: SourceRow[], resources: Resource[], metadata: Metadata): AiUsageItem[] {
  const costs = new Map<string, Decimal>();
  const applications = new Map<string, string>();
  for (const resource of resources) {
    const base = [resource.subscription, resource.resource_group, resource.resource_name].map(normalized).join("\u0000");
    applications.set(base, resource.application || "Unmapped");
    const model = foundryType(resource.service);
    if (model) costs.set(`${base}\u0000${normalized(model)}`, resource.cost);
  }
  const payload = new Map<string, AiUsageItem>();
  for (const row of rows) {
    const base = [row.Subscription, row["Resource Group"], row["Foundry Resource Name"]]
      .map(normalized)
      .join("\u0000");
    const identity = AI_USAGE_COLUMNS.slice(0, 6).map((field) => normalized(row[field])).join("|");
    const usageKey = createHash("sha256").update(identity).digest("hex");
    const counts = {
      input_tokens: count(row["Input Tokens"], "Input Tokens"),
      output_tokens: count(row["Output Tokens"], "Output Tokens"),
      total_tokens: count(row["Total Tokens"], "Total Tokens"),
      requests: count(row.Requests, "Requests"),
    };
    const existing = payload.get(usageKey);
    if (existing) {
      existing.input_tokens += counts.input_tokens;
      existing.output_tokens += counts.output_tokens;
      existing.total_tokens += counts.total_tokens;
      existing.requests += counts.requests;
      continue;
    }
    const cost = costs.get(`${base}\u0000${normalized(row["Model Type"])}`);
    payload.set(usageKey, {
      report_month: monthStart(metadata.period_start),
      billing_period_start: isoDate(metadata.period_start),
      billing_period_end: isoDate(metadata.period_end),
      usage_key: usageKey,
      subscription: text(row.Subscription),
      resource_group: text(row["Resource Group"]),
      foundry_resource_name: text(row["Foundry Resource Name"]),
      application: applications.get(base) ?? null,
      model_deployment_name: text(row["Model Deployment Name"]),
      model_type: text(row["Model Type"]),
      model_version: text(row["Model Version"]) || null,
      ...counts,
      cost_usd: cost?.toString() ?? null,
      cost_scope: cost ? "resource_model" : null,
    });
  }
  return [...payload.values()];
}

export async function transformReportUploads(
  costPaths: string[],
  aiUsagePaths: string[],
): Promise<TransformResult> {
  if (!costPaths.length) throw new Error("At least one cost workbook is required.");
  const costResults = await Promise.all(costPaths.map(readCostWorkbook));
  const resources = combineResources(costResults.flatMap(([items]) => items));
  const metadata = mergeMetadata(costResults.map(([, item]) => item));
  const aiRows: SourceRow[] = [];
  for (const filePath of aiUsagePaths) {
    const parsed = parseCsv(await readFile(filePath, "utf8"), {
      bom: true,
      columns: true,
      skip_empty_lines: true,
    }) as SourceRow[];
    const headers = Object.keys(parsed[0] ?? {});
    const missing = AI_USAGE_COLUMNS.filter((column) => !headers.includes(column));
    if (missing.length) throw new Error(`${path.basename(filePath)}: missing AI usage columns: ${missing.sort().join(", ")}`);
    aiRows.push(...parsed);
  }
  const costItems = costPayload(resources, metadata);
  const aiUsage = aiPayload(aiRows, resources, metadata);
  return {
    report_meta: {
      report_month: monthStart(metadata.period_start),
      period_start: isoDate(metadata.period_start),
      period_end: isoDate(metadata.period_end),
      generated_at: isoDate(metadata.generated_at),
    },
    cost_items: costItems,
    ai_usage: aiUsage,
    stats: {
      cost_rows: costItems.length,
      ai_usage_rows: aiUsage.length,
      total_cost_usd: resources.reduce((sum, resource) => sum.plus(resource.cost), new Decimal(0)).toString(),
      total_tokens: aiUsage.reduce((sum, row) => sum + row.total_tokens, 0),
      requests: aiUsage.reduce((sum, row) => sum + row.requests, 0),
    },
  };
}
