import { createClient, type SupabaseClient } from "@supabase/supabase-js";

import { requiredEnvironment } from "./config.js";
import { ApiError } from "./errors.js";
import type {
  AiUsageFilters,
  AiUsageItem,
  CostFilters,
  CostItem,
  ReportMetadata,
  TransformResult,
} from "./types.js";

const PAGE_SIZE = 1000;
let client: SupabaseClient | undefined;

interface ProjectApplicationInput {
  project_name: string;
  application_name: string;
  description?: string | null;
  is_active?: boolean;
}

interface ResourceMappingInput {
  project_application_id: string;
  subscription_id: string;
  subscription_name?: string | null;
  resource_group_id: string;
  resource_group_name?: string | null;
  resource_id?: string | null;
  resource_name?: string | null;
  environment: string;
  is_active?: boolean;
}

export function getSupabase(): SupabaseClient {
  if (!client) {
    client = createClient(
      requiredEnvironment("SUPABASE_URL"),
      requiredEnvironment("SUPABASE_SERVICE_ROLE_KEY"),
      {
        auth: {
          autoRefreshToken: false,
          detectSessionInUrl: false,
          persistSession: false,
        },
      },
    );
  }
  return client;
}

function assertNoError(
  error: { message: string; code?: string; details?: string } | null,
): void {
  if (!error) return;
  const status =
    error.code === "23505" ? 409 : error.code === "23503" || error.code === "23514" ? 400 : 502;
  throw new ApiError(status, error.message, error.details);
}

async function upsertChunks(
  table: string,
  rows: Record<string, unknown>[],
  onConflict: string,
): Promise<number> {
  const updatedAt = new Date().toISOString();
  for (let start = 0; start < rows.length; start += PAGE_SIZE) {
    const { error } = await getSupabase()
      .from(table)
      .upsert(
        rows.slice(start, start + PAGE_SIZE).map((row) => ({ ...row, updated_at: updatedAt })),
        { onConflict },
      );
    assertNoError(error);
  }
  return rows.length;
}

async function deleteIds(table: string, ids: string[]): Promise<number> {
  for (let start = 0; start < ids.length; start += PAGE_SIZE) {
    const { error } = await getSupabase()
      .from(table)
      .delete()
      .in("id", ids.slice(start, start + PAGE_SIZE));
    assertNoError(error);
  }
  return ids.length;
}

async function removeStaleCostRows(rows: CostItem[]): Promise<number> {
  if (!rows.length) return 0;
  const reportMonth = rows[0].report_month;
  const subscriptionIds = new Set(
    rows.map((row) => row.subscription_id?.trim().toLowerCase()).filter(Boolean),
  );
  const subscriptionNames = new Set(rows.map((row) => row.subscription.trim().toLowerCase()));
  const incoming = new Set(rows.map((row) => `${row.resource_key}\u0000${row.resource_type}`));
  const existing: Array<{
    id: string;
    resource_key: string;
    resource_type: string;
    subscription: string;
    subscription_id: string | null;
  }> = [];

  for (let start = 0; ; start += PAGE_SIZE) {
    const { data, error } = await getSupabase()
      .from("azure_cost_line_items")
      .select("id,resource_key,resource_type,subscription,subscription_id")
      .eq("report_month", reportMonth)
      .range(start, start + PAGE_SIZE - 1);
    assertNoError(error);
    const page = data ?? [];
    existing.push(...page);
    if (page.length < PAGE_SIZE) break;
  }

  const staleIds = existing
    .filter((row) => {
      const id = row.subscription_id?.trim().toLowerCase();
      const inUploadScope = id
        ? subscriptionIds.has(id)
        : subscriptionNames.has(row.subscription.trim().toLowerCase());
      return inUploadScope && !incoming.has(`${row.resource_key}\u0000${row.resource_type}`);
    })
    .map((row) => row.id);
  return deleteIds("azure_cost_line_items", staleIds);
}

async function removeStaleAiRows(rows: AiUsageItem[]): Promise<number> {
  if (!rows.length) return 0;
  const reportMonth = rows[0].report_month;
  const subscriptionNames = new Set(rows.map((row) => row.subscription.trim().toLowerCase()));
  const incoming = new Set(rows.map((row) => row.usage_key));
  const existing: Array<{ id: string; usage_key: string; subscription: string }> = [];

  for (let start = 0; ; start += PAGE_SIZE) {
    const { data, error } = await getSupabase()
      .from("azure_ai_usage")
      .select("id,usage_key,subscription")
      .eq("report_month", reportMonth)
      .range(start, start + PAGE_SIZE - 1);
    assertNoError(error);
    const page = data ?? [];
    existing.push(...page);
    if (page.length < PAGE_SIZE) break;
  }

  const staleIds = existing
    .filter(
      (row) =>
        subscriptionNames.has(row.subscription.trim().toLowerCase()) &&
        !incoming.has(row.usage_key),
    )
    .map((row) => row.id);
  return deleteIds("azure_ai_usage", staleIds);
}

export async function storeTransformedReport(result: TransformResult) {
  const costRows = await upsertChunks(
    "azure_cost_line_items",
    result.cost_items as unknown as Record<string, unknown>[],
    "report_month,resource_key,resource_type",
  );
  const aiRows = await upsertChunks(
    "azure_ai_usage",
    result.ai_usage as unknown as Record<string, unknown>[],
    "report_month,usage_key",
  );
  const removedCostRows = await removeStaleCostRows(result.cost_items);
  const removedAiRows = await removeStaleAiRows(result.ai_usage);
  return {
    cost_rows: costRows,
    ai_usage_rows: aiRows,
    removed_cost_rows: removedCostRows,
    removed_ai_usage_rows: removedAiRows,
  };
}

function applyCostFilters(query: any, filters: CostFilters) {
  if (filters.subscription) query = query.eq("subscription", filters.subscription);
  if (filters.resource_group) query = query.eq("resource_group", filters.resource_group);
  if (filters.project) query = query.eq("resolved_project", filters.project);
  if (filters.application) query = query.eq("resolved_application", filters.application);
  if (filters.environment) query = query.eq("resolved_environment", filters.environment);
  if (filters.resource_type) query = query.eq("resource_type", filters.resource_type);
  return query;
}

function applyAiFilters(query: any, filters: AiUsageFilters) {
  if (filters.subscription) query = query.eq("subscription", filters.subscription);
  if (filters.resource_group) query = query.eq("resource_group", filters.resource_group);
  if (filters.project) query = query.eq("resolved_project", filters.project);
  if (filters.application) query = query.eq("resolved_application", filters.application);
  if (filters.environment) query = query.eq("resolved_environment", filters.environment);
  if (filters.model_type) query = query.eq("model_type", filters.model_type);
  return query;
}

export async function fetchCostItems(
  reportMonth: string,
  filters: CostFilters = {},
  limit?: number,
): Promise<CostItem[]> {
  const rows: CostItem[] = [];
  while (limit === undefined || rows.length < limit) {
    const size = Math.min(PAGE_SIZE, limit === undefined ? PAGE_SIZE : limit - rows.length);
    let query = getSupabase()
      .from("azure_cost_line_items_resolved")
      .select("*")
      .eq("report_month", reportMonth)
      .order("cost_usd", { ascending: false })
      .range(rows.length, rows.length + size - 1);
    query = applyCostFilters(query, filters);
    const { data, error } = await query;
    assertNoError(error);
    const page = (data ?? []) as CostItem[];
    rows.push(...page);
    if (page.length < size) break;
  }
  return limit === undefined ? rows : rows.slice(0, limit);
}

export async function fetchAiUsage(
  reportMonth: string,
  filters: AiUsageFilters = {},
  limit?: number,
): Promise<AiUsageItem[]> {
  const rows: AiUsageItem[] = [];
  while (limit === undefined || rows.length < limit) {
    const size = Math.min(PAGE_SIZE, limit === undefined ? PAGE_SIZE : limit - rows.length);
    let query = getSupabase()
      .from("azure_ai_usage_resolved")
      .select("*")
      .eq("report_month", reportMonth)
      .order("total_tokens", { ascending: false })
      .range(rows.length, rows.length + size - 1);
    query = applyAiFilters(query, filters);
    const { data, error } = await query;
    assertNoError(error);
    const page = (data ?? []) as AiUsageItem[];
    rows.push(...page);
    if (page.length < size) break;
  }
  return limit === undefined ? rows : rows.slice(0, limit);
}

export async function listReportMonths(): Promise<string[]> {
  const { data, error } = await getSupabase()
    .from("azure_report_months")
    .select("report_month")
    .order("report_month", { ascending: false });
  assertNoError(error);
  return [...new Set((data ?? []).map((row) => String(row.report_month)))];
}

export function reportMetadata(rows: CostItem[]): ReportMetadata {
  if (!rows.length) throw new ApiError(404, "No cost data found for this report month.");
  const starts = rows.map((row) => row.billing_period_start).sort();
  const ends = rows.map((row) => row.billing_period_end).sort();
  return {
    report_month: rows[0].report_month,
    period_start: starts[0],
    period_end: ends[ends.length - 1],
    generated_at: new Date().toISOString().slice(0, 10),
  };
}

export function summarizeCosts(rows: CostItem[]) {
  if (!rows.length) throw new ApiError(404, "No cost data found for this report month.");
  const total = rows.reduce((sum, row) => sum + Number(row.cost_usd), 0);
  return {
    report_month: rows[0].report_month,
    billing_period_start: rows[0].billing_period_start,
    billing_period_end: rows[0].billing_period_end,
    currency: rows[0].currency,
    total_cost: total,
    subscription_count: new Set(rows.map((row) => row.subscription_id ?? row.subscription)).size,
    resource_group_count: new Set(rows.map((row) => row.resource_group_id ?? row.resource_group)).size,
    resource_count: new Set(rows.map((row) => row.resource_key)).size,
    line_item_count: rows.length,
    highest_cost_item: rows.reduce((highest, row) =>
      Number(row.cost_usd) > Number(highest.cost_usd) ? row : highest,
    ),
  };
}

async function assertActiveProjectApplication(id: string) {
  const { data, error } = await getSupabase()
    .from("project_applications")
    .select("id,is_active")
    .eq("id", id)
    .maybeSingle();
  assertNoError(error);
  if (!data) throw new ApiError(400, "project_application_id does not exist.");
  if (!data.is_active) {
    throw new ApiError(409, "Mappings cannot be assigned to an inactive application.");
  }
}

export async function listProjectApplications() {
  const { data, error } = await getSupabase()
    .from("project_applications")
    .select("*")
    .order("project_name")
    .order("application_name");
  assertNoError(error);
  return data ?? [];
}

export async function createProjectApplication(input: ProjectApplicationInput) {
  const { data, error } = await getSupabase()
    .from("project_applications")
    .insert({ ...input, is_active: input.is_active ?? true })
    .select("*")
    .single();
  assertNoError(error);
  return data;
}

export async function updateProjectApplication(
  id: string,
  input: Partial<ProjectApplicationInput>,
) {
  const { data, error } = await getSupabase()
    .from("project_applications")
    .update({ ...input, updated_at: new Date().toISOString() })
    .eq("id", id)
    .select("*")
    .single();
  assertNoError(error);
  if (input.is_active === false) {
    const { error: mappingError } = await getSupabase()
      .from("azure_resource_mappings")
      .update({ is_active: false, updated_at: new Date().toISOString() })
      .eq("project_application_id", id)
      .eq("is_active", true);
    assertNoError(mappingError);
  }
  return data;
}

export async function listResourceMappings() {
  const { data, error } = await getSupabase()
    .from("azure_resource_mappings")
    .select("*, project_application:project_applications(*)")
    .order("created_at");
  assertNoError(error);
  return data ?? [];
}

export async function createResourceMapping(input: ResourceMappingInput) {
  await assertActiveProjectApplication(input.project_application_id);
  const { data, error } = await getSupabase()
    .from("azure_resource_mappings")
    .insert({ ...input, is_active: input.is_active ?? true })
    .select("*, project_application:project_applications(*)")
    .single();
  assertNoError(error);
  return data;
}

export async function updateResourceMapping(
  id: string,
  input: Partial<ResourceMappingInput>,
) {
  if (input.project_application_id) {
    await assertActiveProjectApplication(input.project_application_id);
  }
  const { data, error } = await getSupabase()
    .from("azure_resource_mappings")
    .update({ ...input, updated_at: new Date().toISOString() })
    .eq("id", id)
    .select("*, project_application:project_applications(*)")
    .single();
  assertNoError(error);
  return data;
}

export async function archiveResourceMapping(id: string) {
  return updateResourceMapping(id, { is_active: false });
}

export async function resourceCatalog(reportMonth?: string) {
  let query = getSupabase()
    .from("azure_cost_line_items")
    .select(
      "subscription,subscription_id,resource_group,resource_group_id,resource_name,resource_id",
    )
    .not("subscription_id", "is", null)
    .not("resource_group_id", "is", null)
    .order("subscription")
    .order("resource_group")
    .order("resource_name")
    .limit(5000);
  if (reportMonth) query = query.eq("report_month", reportMonth);
  const { data, error } = await query;
  assertNoError(error);

  const subscriptions = new Map<string, any>();
  for (const row of data ?? []) {
    const subscriptionId = String(row.subscription_id);
    const groupId = String(row.resource_group_id);
    if (!subscriptions.has(subscriptionId)) {
      subscriptions.set(subscriptionId, {
        subscription_id: subscriptionId,
        subscription_name: row.subscription,
        resource_groups: new Map<string, any>(),
      });
    }
    const subscription = subscriptions.get(subscriptionId);
    if (!subscription.resource_groups.has(groupId)) {
      subscription.resource_groups.set(groupId, {
        resource_group_id: groupId,
        resource_group_name: row.resource_group,
        resources: new Map<string, any>(),
      });
    }
    const group = subscription.resource_groups.get(groupId);
    if (row.resource_id && !group.resources.has(row.resource_id)) {
      group.resources.set(row.resource_id, {
        resource_id: row.resource_id,
        resource_name: row.resource_name,
      });
    }
  }

  return [...subscriptions.values()].map((subscription) => ({
    ...subscription,
    resource_groups: [...subscription.resource_groups.values()].map((group: any) => ({
      ...group,
      resources: [...group.resources.values()],
    })),
  }));
}

export async function unmappedResources(reportMonth: string) {
  const { data, error } = await getSupabase()
    .from("azure_cost_line_items_resolved")
    .select(
      "subscription,subscription_id,resource_group,resource_group_id,resource_name,resource_id,resource_type,cost_usd",
    )
    .eq("report_month", reportMonth)
    .eq("mapping_level", "unmapped")
    .order("cost_usd", { ascending: false })
    .limit(5000);
  assertNoError(error);
  return data ?? [];
}
