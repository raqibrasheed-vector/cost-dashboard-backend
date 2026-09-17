import assert from "node:assert/strict";
import test from "node:test";

import ExcelJS from "exceljs";

import { buildReportWorkbook } from "../src/report-workbook.js";
import type { AiUsageItem, CostItem, ReportMetadata } from "../src/types.js";

const metadata: ReportMetadata = {
  report_month: "2026-08-01",
  period_start: "2026-08-01",
  period_end: "2026-08-31",
  generated_at: "2026-09-05",
};

const cost: CostItem = {
  report_month: "2026-08-01",
  billing_period_start: "2026-08-01",
  billing_period_end: "2026-08-31",
  currency: "USD",
  resource_key: "resource-1",
  resource_id: "resource-1",
  resource_name: "Foundry Resource",
  resource_group: "Resource Group",
  resource_group_id: "group-1",
  resource_type: "Foundry (gpt-5.1)",
  billing_model: "Usage-Based",
  project: "Project",
  application: "Application",
  environment: "Production",
  subscription: "Subscription",
  subscription_id: "subscription-1",
  location: "eastus",
  cost_usd: "12.345",
};

const usage: AiUsageItem = {
  report_month: "2026-08-01",
  billing_period_start: "2026-08-01",
  billing_period_end: "2026-08-31",
  usage_key: "usage-1",
  subscription: "Subscription",
  resource_group: "Resource Group",
  foundry_resource_name: "Foundry Resource",
  application: "Application",
  model_deployment_name: "deployment",
  model_type: "gpt-5.1",
  model_version: "1",
  input_tokens: 10,
  output_tokens: 5,
  total_tokens: 15,
  requests: 2,
  cost_usd: "12.345",
  cost_scope: "resource_model",
};

test("TypeScript workbook contains the complete report with matching totals", async () => {
  const bytes = await buildReportWorkbook(metadata, [cost], [usage]);
  assert.ok(bytes.length > 10_000);
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(bytes);
  assert.deepEqual(workbook.worksheets.map((sheet) => sheet.name), [
    "Overview",
    "Cost per Application",
    "Cost per Resource Group",
    "Top 10 Resources",
    "Cost per Resource Type",
    "Cost Report",
    "AI Token Usage",
  ]);
  assert.equal(workbook.getWorksheet("Overview")?.getCell("B12").value, 12.34);
  assert.equal(workbook.getWorksheet("Cost Report")?.rowCount, 2);
  assert.equal(workbook.getWorksheet("AI Token Usage")?.rowCount, 2);
});
