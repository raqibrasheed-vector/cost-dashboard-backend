import assert from "node:assert/strict";
import test from "node:test";

import { reportMetadata, summarizeCosts } from "../src/database.js";
import type { CostItem } from "../src/types.js";

function cost(overrides: Partial<CostItem> = {}): CostItem {
  return {
    report_month: "2026-08-01",
    billing_period_start: "2026-08-01",
    billing_period_end: "2026-08-31",
    currency: "USD",
    resource_key: "resource-1",
    resource_id: "resource-1",
    resource_name: "Resource 1",
    resource_group: "Group 1",
    resource_group_id: "group-1",
    resource_type: "Azure OpenAI",
    billing_model: "Consumption",
    project: "Cirrus",
    application: "Reporting",
    environment: "Production",
    subscription: "Subscription 1",
    subscription_id: "subscription-1",
    location: "eastus",
    cost_usd: "10.25",
    ...overrides,
  };
}

test("reportMetadata spans all supplied rows", () => {
  const metadata = reportMetadata([
    cost({ billing_period_start: "2026-08-03", billing_period_end: "2026-08-20" }),
    cost({ billing_period_start: "2026-08-01", billing_period_end: "2026-08-31" }),
  ]);
  assert.equal(metadata.report_month, "2026-08-01");
  assert.equal(metadata.period_start, "2026-08-01");
  assert.equal(metadata.period_end, "2026-08-31");
});

test("summarizeCosts calculates totals and distinct hierarchy counts", () => {
  const summary = summarizeCosts([
    cost(),
    cost({
      resource_key: "resource-2",
      resource_id: "resource-2",
      resource_name: "Resource 2",
      resource_group: "Group 2",
      resource_group_id: "group-2",
      cost_usd: "4.75",
    }),
  ]);
  assert.equal(summary.total_cost, 15);
  assert.equal(summary.subscription_count, 1);
  assert.equal(summary.resource_group_count, 2);
  assert.equal(summary.resource_count, 2);
  assert.equal(summary.highest_cost_item.resource_key, "resource-1");
});

test("report helpers reject empty data", () => {
  assert.throws(() => reportMetadata([]), /No cost data/);
  assert.throws(() => summarizeCosts([]), /No cost data/);
});
