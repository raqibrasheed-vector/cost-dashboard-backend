import assert from "node:assert/strict";
import { randomUUID } from "node:crypto";
import { basename } from "node:path";
import { readFile } from "node:fs/promises";
import type { AddressInfo } from "node:net";
import type { Server } from "node:http";

import { createClient } from "@supabase/supabase-js";
import dotenv from "dotenv";

dotenv.config({ quiet: true });
process.env.NODE_ENV = "test";

const [costPath, aiPath] = process.argv.slice(2);
if (!costPath || !aiPath) {
  throw new Error("Usage: npm run test:live -- <cost-report.xlsx> <ai-usage.csv>");
}

const supabaseUrl = process.env.SUPABASE_URL;
const supabaseKey = process.env.SUPABASE_SERVICE_ROLE_KEY;
if (!supabaseUrl || !supabaseKey) throw new Error("Supabase environment variables are required.");

const supabase = createClient(supabaseUrl, supabaseKey, {
  auth: { autoRefreshToken: false, detectSessionInUrl: false, persistSession: false },
});
const createdUserIds: string[] = [];
const createdApplicationIds: string[] = [];
const createdMappingIds: string[] = [];
let server: Server | undefined;

function check(error: { message: string } | null): void {
  if (error) throw new Error(error.message);
}

async function createTestUser(role?: "admin"): Promise<string> {
  const suffix = randomUUID();
  const email = `cost-report-e2e-${suffix}@example.com`;
  const password = `Aa1!${randomUUID()}`;
  const { data, error } = await supabase.auth.admin.createUser({
    email,
    password,
    email_confirm: true,
    app_metadata: role ? { role } : {},
  });
  check(error);
  assert.ok(data.user);
  createdUserIds.push(data.user.id);

  const authClient = createClient(supabaseUrl!, supabaseKey!, {
    auth: { autoRefreshToken: false, detectSessionInUrl: false, persistSession: false },
  });
  const { data: session, error: signInError } = await authClient.auth.signInWithPassword({
    email,
    password,
  });
  check(signInError);
  assert.ok(session.session?.access_token);
  return session.session.access_token;
}

async function body(response: Response): Promise<any> {
  const payload = await response.json();
  if (!response.ok) throw new Error(`${response.status}: ${JSON.stringify(payload)}`);
  return payload;
}

async function main() {
  const adminToken = await createTestUser("admin");
  const userToken = await createTestUser();
  const { createApp } = await import("../src/server.js");
  const app = await createApp();
  server = await new Promise<Server>((resolve) => {
    const listening = app.listen(0, "127.0.0.1", () => resolve(listening));
  });
  const port = (server.address() as AddressInfo).port;
  const baseUrl = `http://127.0.0.1:${port}`;

  const request = (path: string, token?: string, init: RequestInit = {}) => {
    const headers = new Headers(init.headers);
    if (token) headers.set("authorization", `Bearer ${token}`);
    return fetch(`${baseUrl}${path}`, { ...init, headers });
  };
  const jsonRequest = (path: string, token: string, method: string, payload: unknown) =>
    request(path, token, {
      method,
      headers: { "content-type": "application/json" },
      body: JSON.stringify(payload),
    });

  assert.equal((await request("/health")).status, 200);
  assert.equal((await request("/docs/")).status, 200);
  const openApi = await body(await request("/openapi.json"));
  assert.equal(openApi.openapi, "3.0.3");
  assert.deepEqual(
    Object.keys(openApi.paths).sort(),
    [
      "/api/ai/usage",
      "/api/costs",
      "/api/overview/summary",
      "/api/reports",
      "/api/reports/ingest",
      "/api/reports/{reportMonth}/excel",
      "/api/settings/project-applications",
      "/api/settings/project-applications/{id}",
      "/api/settings/resource-catalog",
      "/api/settings/resource-mappings",
      "/api/settings/resource-mappings/{id}",
      "/api/settings/unmapped",
      "/health",
      "/openapi.json",
    ].sort(),
  );
  assert.equal((await request("/api/reports")).status, 401);
  assert.equal((await request("/api/reports", userToken)).status, 200);
  assert.equal((await request("/api/not-a-route", userToken)).status, 404);
  assert.equal((await request("/api/costs?report_month=2026-08-02", userToken)).status, 400);
  assert.equal(
    (
      await jsonRequest("/api/settings/project-applications", userToken, "POST", {
        project_name: "Forbidden",
        application_name: "Forbidden",
      })
    ).status,
    403,
  );

  const { data: seedCost, error: seedCostError } = await supabase
    .from("azure_cost_line_items")
    .select("*")
    .eq("report_month", "2026-08-01")
    .limit(1)
    .single();
  check(seedCostError);
  const { id: _costId, created_at: _costCreated, updated_at: _costUpdated, ...costFields } = seedCost;
  const staleCostKey = `e2e-stale-${randomUUID()}`;
  const { error: staleCostError } = await supabase.from("azure_cost_line_items").insert({
    ...costFields,
    resource_key: staleCostKey,
    resource_id: staleCostKey,
    resource_name: "E2E stale resource",
    resource_type: "E2E stale type",
    cost_usd: "0.01",
  });
  check(staleCostError);

  const { data: seedAi, error: seedAiError } = await supabase
    .from("azure_ai_usage")
    .select("*")
    .eq("report_month", "2026-08-01")
    .limit(1)
    .single();
  check(seedAiError);
  const { id: _aiId, created_at: _aiCreated, updated_at: _aiUpdated, ...aiFields } = seedAi;
  const staleUsageKey = `e2e-stale-${randomUUID()}`;
  const { error: staleAiError } = await supabase.from("azure_ai_usage").insert({
    ...aiFields,
    usage_key: staleUsageKey,
    model_deployment_name: "E2E stale deployment",
    input_tokens: 1,
    output_tokens: 1,
    total_tokens: 2,
    requests: 1,
  });
  check(staleAiError);

  const form = new FormData();
  form.append("cost_reports", new Blob([await readFile(costPath)]), basename(costPath));
  form.append("ai_usage", new Blob([await readFile(aiPath)]), basename(aiPath));
  const ingestResponse = await request("/api/reports/ingest", adminToken, {
    method: "POST",
    body: form,
  });
  assert.equal(ingestResponse.status, 201);
  const ingest = await body(ingestResponse);
  assert.equal(ingest.stats.cost_rows, 214);
  assert.equal(ingest.stats.ai_usage_rows, 18);
  assert.equal(ingest.stored.removed_cost_rows, 1);
  assert.equal(ingest.stored.removed_ai_usage_rows, 1);

  const costs = await body(
    await request("/api/costs?report_month=2026-08-01&limit=5000", userToken),
  );
  const aiUsage = await body(
    await request("/api/ai/usage?report_month=2026-08-01&limit=5000", userToken),
  );
  assert.equal(costs.count, 214);
  assert.equal(aiUsage.count, 18);

  const firstCost = costs.items[0];
  for (const [parameter, field] of [
    ["subscription", "subscription"],
    ["resource_group", "resource_group"],
    ["resource_type", "resource_type"],
  ]) {
    const query = new URLSearchParams({
      report_month: "2026-08-01",
      limit: "5000",
      [parameter]: firstCost[field],
    });
    const filtered = await body(await request(`/api/costs?${query}`, userToken));
    assert.ok(filtered.count > 0);
    assert.ok(filtered.items.every((item: any) => item[field] === firstCost[field]));
  }

  const firstAi = aiUsage.items[0];
  const aiQuery = new URLSearchParams({
    report_month: "2026-08-01",
    limit: "5000",
    model_type: firstAi.model_type,
  });
  const filteredAi = await body(await request(`/api/ai/usage?${aiQuery}`, userToken));
  assert.ok(filteredAi.count > 0);
  assert.ok(filteredAi.items.every((item: any) => item.model_type === firstAi.model_type));

  const summary = await body(
    await request("/api/overview/summary?report_month=2026-08-01", userToken),
  );
  assert.equal(summary.line_item_count, 214);
  assert.ok(Math.abs(summary.total_cost - 5786.764892263762) < 0.000001);

  const workbook = await request("/api/reports/2026-08-01/excel", userToken);
  assert.equal(workbook.status, 200);
  assert.match(workbook.headers.get("content-type") ?? "", /spreadsheetml/);
  const workbookBytes = new Uint8Array(await workbook.arrayBuffer());
  assert.ok(workbookBytes.length > 10_000);
  assert.equal(String.fromCharCode(workbookBytes[0], workbookBytes[1]), "PK");

  const catalog = await body(
    await request("/api/settings/resource-catalog?report_month=2026-08-01", userToken),
  );
  const unmapped = await body(
    await request("/api/settings/unmapped?report_month=2026-08-01", userToken),
  );
  assert.ok(Array.isArray(unmapped.items));
  const mappings = await body(await request("/api/settings/resource-mappings", userToken));
  const activeGroupKeys = new Set(
    mappings.items
      .filter((item: any) => item.is_active && !item.resource_id)
      .map((item: any) => `${item.subscription_id.toLowerCase()}|${item.resource_group_id.toLowerCase()}`),
  );
  const activeResourceKeys = new Set(
    mappings.items
      .filter((item: any) => item.is_active && item.resource_id)
      .map((item: any) => `${item.subscription_id.toLowerCase()}|${item.resource_id.toLowerCase()}`),
  );
  let selected: any;
  for (const subscription of catalog.items) {
    for (const group of subscription.resource_groups) {
      const groupKey = `${subscription.subscription_id.toLowerCase()}|${group.resource_group_id.toLowerCase()}`;
      const resource = group.resources.find(
        (item: any) =>
          !activeResourceKeys.has(
            `${subscription.subscription_id.toLowerCase()}|${item.resource_id.toLowerCase()}`,
          ),
      );
      if (!activeGroupKeys.has(groupKey) && resource) {
        selected = { subscription, group, resource };
        break;
      }
    }
    if (selected) break;
  }
  assert.ok(selected, "No unmapped resource group was available for the mapping test.");

  const suffix = randomUUID().slice(0, 8);
  const groupApp = await body(
    await jsonRequest("/api/settings/project-applications", adminToken, "POST", {
      project_name: `E2E Project ${suffix}`,
      application_name: `Group App ${suffix}`,
    }),
  );
  createdApplicationIds.push(groupApp.id);
  const resourceApp = await body(
    await jsonRequest("/api/settings/project-applications", adminToken, "POST", {
      project_name: `E2E Project ${suffix}`,
      application_name: `Resource App ${suffix}`,
    }),
  );
  createdApplicationIds.push(resourceApp.id);
  const applications = await body(
    await request("/api/settings/project-applications", userToken),
  );
  assert.ok(applications.items.some((item: any) => item.id === groupApp.id));
  assert.ok(applications.items.some((item: any) => item.id === resourceApp.id));

  const mappingBase = {
    subscription_id: selected.subscription.subscription_id,
    subscription_name: selected.subscription.subscription_name,
    resource_group_id: selected.group.resource_group_id,
    resource_group_name: selected.group.resource_group_name,
    environment: "E2E",
  };
  const groupMapping = await body(
    await jsonRequest("/api/settings/resource-mappings", adminToken, "POST", {
      ...mappingBase,
      project_application_id: groupApp.id,
    }),
  );
  createdMappingIds.push(groupMapping.id);

  let resolved = await body(
    await request("/api/costs?report_month=2026-08-01&limit=5000", userToken),
  );
  let selectedRows = resolved.items.filter(
    (item: any) =>
      item.subscription_id === selected.subscription.subscription_id &&
      item.resource_group_id === selected.group.resource_group_id,
  );
  assert.ok(selectedRows.length > 0);
  assert.ok(selectedRows.every((item: any) => item.resolved_application === groupApp.application_name));
  assert.ok(selectedRows.every((item: any) => item.mapping_level === "resource_group"));

  const resourceMapping = await body(
    await jsonRequest("/api/settings/resource-mappings", adminToken, "POST", {
      ...mappingBase,
      project_application_id: resourceApp.id,
      resource_id: selected.resource.resource_id,
      resource_name: selected.resource.resource_name,
    }),
  );
  createdMappingIds.push(resourceMapping.id);

  resolved = await body(
    await request("/api/costs?report_month=2026-08-01&limit=5000", userToken),
  );
  let selectedResourceRows = resolved.items.filter(
    (item: any) => item.resource_id === selected.resource.resource_id,
  );
  assert.ok(selectedResourceRows.length > 0);
  assert.ok(
    selectedResourceRows.every(
      (item: any) =>
        item.resolved_application === resourceApp.application_name &&
        item.mapping_level === "resource",
    ),
  );

  const updatedResourceMapping = await body(
    await jsonRequest(
      `/api/settings/resource-mappings/${resourceMapping.id}`,
      adminToken,
      "PATCH",
      { environment: "E2E Updated" },
    ),
  );
  assert.equal(updatedResourceMapping.environment, "E2E Updated");

  assert.equal(
    (await request(`/api/settings/resource-mappings/${resourceMapping.id}`, adminToken, { method: "DELETE" }))
      .status,
    200,
  );
  resolved = await body(
    await request("/api/costs?report_month=2026-08-01&limit=5000", userToken),
  );
  selectedResourceRows = resolved.items.filter(
    (item: any) => item.resource_id === selected.resource.resource_id,
  );
  assert.ok(
    selectedResourceRows.every(
      (item: any) =>
        item.resolved_application === groupApp.application_name &&
        item.mapping_level === "resource_group",
    ),
  );

  assert.equal(
    (
      await jsonRequest(
        `/api/settings/project-applications/${groupApp.id}`,
        adminToken,
        "PATCH",
        { is_active: false },
      )
    ).status,
    200,
  );
  const afterDeactivate = await body(
    await request("/api/settings/resource-mappings", userToken),
  );
  assert.equal(afterDeactivate.items.find((item: any) => item.id === groupMapping.id).is_active, false);
  assert.equal(
    (
      await jsonRequest("/api/settings/resource-mappings", adminToken, "POST", {
        ...mappingBase,
        project_application_id: groupApp.id,
      })
    ).status,
    409,
  );

  console.log(
    JSON.stringify({
      authentication: "passed",
      swagger_contract: "passed",
      route_matrix: "passed",
      multipart_ingestion: "passed",
      stale_row_replacement: "passed",
      cost_rows: costs.count,
      ai_usage_rows: aiUsage.count,
      filters: "passed",
      excel_download_bytes: workbookBytes.length,
      mapping_precedence: "passed",
      application_deactivation: "passed",
    }),
  );
}

try {
  await main();
} finally {
  if (server) await new Promise<void>((resolve, reject) => server!.close((error) => (error ? reject(error) : resolve())));
  if (createdMappingIds.length) {
    const { error } = await supabase.from("azure_resource_mappings").delete().in("id", createdMappingIds);
    check(error);
  }
  if (createdApplicationIds.length) {
    const { error } = await supabase.from("project_applications").delete().in("id", createdApplicationIds);
    check(error);
  }
  for (const id of createdUserIds) {
    const { error } = await supabase.auth.admin.deleteUser(id);
    check(error);
  }
}
