# Cost Report API and Data Guide

## Connection and Authentication

Local base URL: `http://127.0.0.1:3001`

- `GET /health`, `GET /openapi.json`, and `/docs` are public.
- Every `/api/*` route requires `Authorization: Bearer <supabase-user-jwt>`.
- `POST`, `PATCH`, and `DELETE` operations require the Supabase user to have
  `app_metadata.role = "admin"`.
- The Supabase service-role key belongs only in the backend `.env`. It must
  never be sent to or stored by the frontend.

Interactive API documentation is available at `/docs`.

## Verifying Every API

Run the live route matrix with representative source files:

```powershell
npm run test:live -- "path\to\cost-report.xlsx" "path\to\ai-usage.csv"
```

This verifies authentication and authorization, Swagger/OpenAPI, report
ingestion, stale-row replacement, cost and AI reads, filters, summary, Excel
download, settings reads and mutations, mapping precedence, deactivation,
validation errors, and unknown-route handling. Temporary users and settings
records are removed when the run finishes.

## Report APIs

| Method | Route | Purpose |
| --- | --- | --- |
| `POST` | `/api/reports/ingest` | Parse uploaded files and store the transformed data. Admin only. |
| `GET` | `/api/reports` | List report months available in the database. |
| `GET` | `/api/costs` | Return filtered, resolved Azure cost rows. |
| `GET` | `/api/ai/usage` | Return filtered AI model usage rows. |
| `GET` | `/api/overview/summary` | Return report totals and hierarchy counts. |
| `GET` | `/api/reports/:reportMonth/excel` | Generate and download a filtered seven-sheet workbook. |

### Upload Contract

`POST /api/reports/ingest` uses `multipart/form-data`:

- `cost_reports`: one to 20 `.xlsx` files; required.
- `ai_usage`: zero to 20 `.csv` files; optional.
- Maximum size is 50 MB per file.
- All cost workbooks in one request must cover the same report month.

Files are combined before ingestion. Duplicate resource/model identities are
aggregated, so uploading the same source file twice would double its values.
Different report months must be sent as separate requests. Invalid input
returns `422` before database ingestion, and temporary files are deleted after
every request.

### Query Parameters

`report_month` is required where shown and uses the first day of the month,
for example `2026-08-01`.

Cost, summary, and Excel routes support exact-match filters:

- `subscription`
- `resource_group`
- `project`
- `application`
- `environment`
- `resource_type` (cost and Excel)
- `limit` from 1 to 5000 (JSON row endpoints)

AI usage supports the same hierarchy filters plus `model_type` instead of
`resource_type`.

## Settings APIs

| Method | Route | Purpose |
| --- | --- | --- |
| `GET` | `/api/settings/project-applications` | List project/application records. |
| `POST` | `/api/settings/project-applications` | Create a project/application record. Admin only. |
| `PATCH` | `/api/settings/project-applications/:id` | Update or deactivate a record. Admin only. |
| `GET` | `/api/settings/resource-mappings` | List mappings with their project/application details. |
| `POST` | `/api/settings/resource-mappings` | Map a resource group or individual resource. Admin only. |
| `PATCH` | `/api/settings/resource-mappings/:id` | Update a mapping. Admin only. |
| `DELETE` | `/api/settings/resource-mappings/:id` | Archive a mapping by setting it inactive. Admin only. |
| `GET` | `/api/settings/resource-catalog` | Return the discovered subscription/resource hierarchy. Optional `report_month`. |
| `GET` | `/api/settings/unmapped` | Return unmapped costs for a required `report_month`. |

Project/application input fields:

```json
{
  "project_name": "Knowledge Assistant",
  "application_name": "OCM20 Evolve",
  "description": "Optional",
  "is_active": true
}
```

Resource mapping input fields:

```json
{
  "project_application_id": "uuid",
  "subscription_id": "azure-subscription-id",
  "subscription_name": "Optional display name",
  "resource_group_id": "azure-resource-group-id",
  "resource_group_name": "Optional display name",
  "resource_id": null,
  "resource_name": null,
  "environment": "Production",
  "is_active": true
}
```

Leave both `resource_id` and `resource_name` null for a resource-group mapping.
Supply both for a resource-specific override. Resolution order is:

1. Active resource-specific mapping
2. Active resource-group mapping
3. Values derived while ingesting the source report
4. `Unmapped`

Deactivating a project/application also deactivates its active mappings.

## Stored Data

The backend stores transformed records, not Excel row numbers, pivot tables,
generated workbooks, or permanent copies of uploaded files.

### `azure_cost_line_items`

One row per report month, scoped Azure resource identity, and effective
resource/model type. Unique identity:
`report_month + resource_key + resource_type`.

Stored fields:

- Billing: `report_month`, `billing_period_start`, `billing_period_end`,
  `currency`, `cost_usd`, `billing_model`
- Subscription: `subscription`, `subscription_id`
- Resource group: `resource_group`, `resource_group_id`
- Resource: `resource_key`, `resource_id`, `resource_name`, `resource_type`,
  `location`
- Ingestion-time classification: `project`, `application`, `environment`
- Database metadata: generated `id`, `created_at`, `updated_at`

`resource_key` uses the Azure resource ID when available. Otherwise it is
derived from subscription, resource group, and resource name.

### `azure_ai_usage`

One row per report month and AI deployment/model identity. Unique identity:
`report_month + usage_key`.

Stored fields:

- Period: `report_month`, `billing_period_start`, `billing_period_end`
- Scope: `subscription`, `resource_group`, `foundry_resource_name`,
  `application`
- Model: `model_deployment_name`, `model_type`, `model_version`
- Usage: `input_tokens`, `output_tokens`, `total_tokens`, `requests`
- Cost matching: `cost_usd`, `cost_scope`
- Database metadata: generated `id`, `created_at`, `updated_at`

`usage_key` is a stable hash of the source identity columns. AI cost is matched
at resource/model scope and must not be counted once per deployment.

### `project_applications`

Stores configurable project/application pairs:

- `id`
- `project_name`
- `application_name`
- `description`
- `is_active`
- `created_at`, `updated_at`

### `azure_resource_mappings`

Stores application and environment assignments at resource-group or resource
scope:

- `id`, `project_application_id`
- `subscription_id`, `subscription_name`
- `resource_group_id`, `resource_group_name`
- Optional `resource_id`, `resource_name`
- `environment`, `is_active`
- `created_at`, `updated_at`

The API reads resolved database views so frontend responses additionally
contain `resolved_project`, `resolved_application`, `resolved_environment`,
`mapping_level`, and `mapping_id` without rewriting original cost records.

## Generated Excel Workbook

Excel downloads are built on demand from filtered database rows and streamed
directly to the caller. The workbook contains:

1. `Overview`
2. `Cost per Application`
3. `Cost per Resource Group`
4. `Top 10 Resources`
5. `Cost per Resource Type`
6. `Cost Report`
7. `AI Token Usage`
