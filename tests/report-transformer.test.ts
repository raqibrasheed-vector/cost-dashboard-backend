import assert from "node:assert/strict";
import { mkdtemp, rm, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";

import ExcelJS from "exceljs";

import { transformReportUploads } from "../src/report-transformer.js";

test("TypeScript transformer aggregates Azure cost rows and matches AI usage", async () => {
  const directory = await mkdtemp(path.join(os.tmpdir(), "cost-report-transform-"));
  try {
    const costPath = path.join(directory, "cost.xlsx");
    const aiPath = path.join(directory, "usage.csv");
    const workbook = new ExcelJS.Workbook();
    const summary = workbook.addWorksheet("Summary");
    summary.addRows([
      ["Start date:", "Aug 1, 2026"],
      ["End date:", "Aug 31, 2026"],
      ["Generated:", "Sep 5, 2026"],
    ]);
    const data = workbook.addWorksheet("Data");
    data.addRow([
      "Resource",
      "ResourceId",
      "ResourceType",
      "ResourceGroupName",
      "ResourceGroupId",
      "ResourceLocation",
      "SubscriptionName",
      "SubscriptionId",
      "ServiceName",
      "ServiceTier",
      "Product",
      "Meter",
      "Cost",
      "CostUSD",
    ]);
    for (const cost of ["1.10", "2.20"]) {
      data.addRow([
        "aif-test",
        "/subscriptions/sub/resourceGroups/rg/providers/Microsoft.CognitiveServices/accounts/aif-test",
        "Cognitive Services account",
        "rg-ocm20_evolve-dev-001",
        "/subscriptions/sub/resourceGroups/rg-ocm20_evolve-dev-001",
        "eastus2",
        "Test Subscription",
        "sub",
        "Foundry Models",
        "Azure OpenAI",
        "Azure OpenAI GPT 5.1",
        "GPT 5.1 tokens",
        cost,
        cost,
      ]);
    }
    await workbook.xlsx.writeFile(costPath);
    await writeFile(
      aiPath,
      [
        "Subscription,Foundry Resource Name,Resource Group,Model Deployment Name,Model Type,Model Version,Input Tokens,Output Tokens,Total Tokens,Requests",
        "Test Subscription,aif-test,rg-ocm20_evolve-dev-001,gpt-test,gpt-5.1,1,10,5,15,2",
      ].join("\n"),
    );

    const result = await transformReportUploads([costPath], [aiPath]);
    assert.equal(result.report_meta.report_month, "2026-08-01");
    assert.equal(result.cost_items.length, 1);
    assert.equal(result.cost_items[0].cost_usd, "3.3");
    assert.equal(result.cost_items[0].resource_type, "Foundry (gpt-5.1)");
    assert.equal(result.cost_items[0].application, "OCM20 Evolve (EA Sandbox)");
    assert.equal(result.ai_usage.length, 1);
    assert.equal(result.ai_usage[0].cost_usd, "3.3");
    assert.equal(result.stats.total_tokens, 15);
  } finally {
    await rm(directory, { recursive: true, force: true });
  }
});
