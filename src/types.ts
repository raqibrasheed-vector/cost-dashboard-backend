export interface ReportMetadata {
  report_month: string;
  period_start: string;
  period_end: string;
  generated_at: string;
}

export interface CostItem {
  report_month: string;
  billing_period_start: string;
  billing_period_end: string;
  currency: string;
  resource_key: string;
  resource_id: string | null;
  resource_name: string | null;
  resource_group: string;
  resource_group_id: string | null;
  resource_type: string;
  billing_model: string;
  project: string;
  application: string;
  environment: string;
  subscription: string;
  subscription_id: string | null;
  location: string | null;
  cost_usd: string | number;
  resolved_project?: string;
  resolved_application?: string;
  resolved_environment?: string;
  mapping_level?: "resource" | "resource_group" | "unmapped";
  mapping_id?: string | null;
}

export interface AiUsageItem {
  report_month: string;
  billing_period_start: string;
  billing_period_end: string;
  usage_key: string;
  subscription: string;
  resource_group: string;
  foundry_resource_name: string;
  application: string | null;
  model_deployment_name: string;
  model_type: string;
  model_version: string | null;
  input_tokens: number;
  output_tokens: number;
  total_tokens: number;
  requests: number;
  cost_usd: string | number | null;
  cost_scope: "resource_model" | null;
  resolved_project?: string;
  resolved_application?: string;
  resolved_environment?: string;
  mapping_level?: "resource" | "resource_group" | "unmapped";
  mapping_id?: string | null;
}

export interface TransformResult {
  report_meta: ReportMetadata;
  cost_items: CostItem[];
  ai_usage: AiUsageItem[];
  stats: {
    cost_rows: number;
    ai_usage_rows: number;
    total_cost_usd: string;
    total_tokens: number;
    requests: number;
  };
}

export interface CostFilters {
  subscription?: string;
  resource_group?: string;
  project?: string;
  application?: string;
  environment?: string;
  resource_type?: string;
}

export interface AiUsageFilters {
  subscription?: string;
  resource_group?: string;
  project?: string;
  application?: string;
  environment?: string;
  model_type?: string;
}
