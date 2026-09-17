const bearerSecurity = [{ bearerAuth: [] }];
const json = (schema: object) => ({ "application/json": { schema } });
const response = (description: string, schema: object) => ({
  description,
  content: json(schema),
});
const errorResponse = (description: string) =>
  response(description, { $ref: "#/components/schemas/Error" });

const reportMonthQuery = {
  name: "report_month",
  in: "query",
  required: true,
  description: "First day of the report month.",
  schema: { type: "string", format: "date", example: "2026-08-01" },
};
const reportMonthPath = {
  name: "reportMonth",
  in: "path",
  required: true,
  schema: { type: "string", format: "date", example: "2026-08-01" },
};
const idPath = {
  name: "id",
  in: "path",
  required: true,
  schema: { type: "string", format: "uuid" },
};
const limitQuery = {
  name: "limit",
  in: "query",
  schema: { type: "integer", minimum: 1, maximum: 5000, default: 1000 },
};
const queryString = (name: string) => ({
  name,
  in: "query",
  description: `Exact-match ${name.replaceAll("_", " ")} filter.`,
  schema: { type: "string" },
});
const hierarchyFilters = [
  "subscription",
  "resource_group",
  "project",
  "application",
  "environment",
].map(queryString);
const costFilters = [...hierarchyFilters, queryString("resource_type")];
const aiFilters = [...hierarchyFilters, queryString("model_type")];

const projectApplicationInput = {
  type: "object",
  additionalProperties: false,
  required: ["project_name", "application_name"],
  properties: {
    project_name: { type: "string", minLength: 1 },
    application_name: { type: "string", minLength: 1 },
    description: { type: "string", nullable: true },
    is_active: { type: "boolean", default: true },
  },
};
const resourceMappingInput = {
  type: "object",
  additionalProperties: false,
  required: ["project_application_id", "subscription_id", "resource_group_id", "environment"],
  properties: {
    project_application_id: { type: "string", format: "uuid" },
    subscription_id: { type: "string", minLength: 1 },
    subscription_name: { type: "string", nullable: true },
    resource_group_id: { type: "string", minLength: 1 },
    resource_group_name: { type: "string", nullable: true },
    resource_id: {
      type: "string",
      nullable: true,
      description: "Supply together with resource_name for a resource override.",
    },
    resource_name: { type: "string", nullable: true },
    environment: { type: "string", minLength: 1 },
    is_active: { type: "boolean", default: true },
  },
};

export const openApiDocument = {
  openapi: "3.0.3",
  info: {
    title: "Azure Cost Report API",
    version: "1.0.0",
    description: "Node.js API for cost ingestion, Supabase reporting, mappings, and Excel export.",
  },
  servers: [{ url: "http://127.0.0.1:3001", description: "Local development" }],
  security: bearerSecurity,
  components: {
    securitySchemes: {
      bearerAuth: {
        type: "http",
        scheme: "bearer",
        bearerFormat: "Supabase user JWT",
        description: "Mutations require app_metadata.role=admin.",
      },
    },
    schemas: {
      Error: {
        type: "object",
        required: ["error"],
        properties: { error: { type: "string" }, details: { nullable: true } },
      },
      ReportMetadata: {
        type: "object",
        properties: {
          report_month: { type: "string", format: "date" },
          period_start: { type: "string", format: "date" },
          period_end: { type: "string", format: "date" },
          generated_at: { type: "string", format: "date" },
        },
      },
      CostItem: {
        type: "object",
        properties: {
          id: { type: "string", format: "uuid" },
          report_month: { type: "string", format: "date" },
          billing_period_start: { type: "string", format: "date" },
          billing_period_end: { type: "string", format: "date" },
          currency: { type: "string", example: "USD" },
          resource_key: { type: "string" },
          resource_id: { type: "string", nullable: true },
          resource_name: { type: "string", nullable: true },
          resource_group: { type: "string" },
          resource_group_id: { type: "string", nullable: true },
          resource_type: { type: "string" },
          billing_model: { type: "string" },
          project: { type: "string" },
          application: { type: "string" },
          environment: { type: "string" },
          subscription: { type: "string" },
          subscription_id: { type: "string", nullable: true },
          location: { type: "string", nullable: true },
          cost_usd: { oneOf: [{ type: "number" }, { type: "string" }] },
          resolved_project: { type: "string" },
          resolved_application: { type: "string" },
          resolved_environment: { type: "string" },
          mapping_level: { type: "string", enum: ["resource", "resource_group", "unmapped"] },
          mapping_id: { type: "string", format: "uuid", nullable: true },
        },
      },
      AiUsageItem: {
        type: "object",
        properties: {
          id: { type: "string", format: "uuid" },
          report_month: { type: "string", format: "date" },
          billing_period_start: { type: "string", format: "date" },
          billing_period_end: { type: "string", format: "date" },
          usage_key: { type: "string" },
          subscription: { type: "string" },
          resource_group: { type: "string" },
          foundry_resource_name: { type: "string" },
          application: { type: "string", nullable: true },
          model_deployment_name: { type: "string" },
          model_type: { type: "string" },
          model_version: { type: "string", nullable: true },
          input_tokens: { type: "integer", format: "int64" },
          output_tokens: { type: "integer", format: "int64" },
          total_tokens: { type: "integer", format: "int64" },
          requests: { type: "integer", format: "int64" },
          cost_usd: { oneOf: [{ type: "number" }, { type: "string" }], nullable: true },
          cost_scope: { type: "string", enum: ["resource_model"], nullable: true },
          resolved_project: { type: "string" },
          resolved_application: { type: "string" },
          resolved_environment: { type: "string" },
          mapping_level: { type: "string", enum: ["resource", "resource_group", "unmapped"] },
          mapping_id: { type: "string", format: "uuid", nullable: true },
        },
      },
      ProjectApplication: {
        type: "object",
        properties: {
          id: { type: "string", format: "uuid" },
          project_name: { type: "string" },
          application_name: { type: "string" },
          description: { type: "string", nullable: true },
          is_active: { type: "boolean" },
          created_at: { type: "string", format: "date-time" },
          updated_at: { type: "string", format: "date-time" },
        },
      },
      ResourceMapping: {
        type: "object",
        properties: {
          id: { type: "string", format: "uuid" },
          project_application_id: { type: "string", format: "uuid" },
          subscription_id: { type: "string" },
          subscription_name: { type: "string", nullable: true },
          resource_group_id: { type: "string" },
          resource_group_name: { type: "string", nullable: true },
          resource_id: { type: "string", nullable: true },
          resource_name: { type: "string", nullable: true },
          environment: { type: "string" },
          is_active: { type: "boolean" },
          project_application: { $ref: "#/components/schemas/ProjectApplication" },
        },
      },
    },
  },
  paths: {
    "/health": {
      get: {
        operationId: "getHealth",
        summary: "Service health",
        security: [],
        responses: {
          "200": response("Healthy", {
            type: "object",
            properties: {
              status: { type: "string", example: "ok" },
              runtime: { type: "string", example: "node" },
              report_engine: { type: "string", example: "typescript" },
            },
          }),
        },
      },
    },
    "/openapi.json": {
      get: {
        operationId: "getOpenApiDocument",
        summary: "OpenAPI document",
        security: [],
        responses: { "200": response("OpenAPI 3.0 document", { type: "object" }) },
      },
    },
    "/api/reports/ingest": {
      post: {
        operationId: "ingestReports",
        summary: "Transform and store uploaded cost and AI usage files",
        description: "Admin only. All cost workbooks in one request must cover the same month.",
        requestBody: {
          required: true,
          content: {
            "multipart/form-data": {
              schema: {
                type: "object",
                required: ["cost_reports"],
                properties: {
                  cost_reports: {
                    type: "array",
                    minItems: 1,
                    maxItems: 20,
                    items: { type: "string", format: "binary" },
                  },
                  ai_usage: {
                    type: "array",
                    maxItems: 20,
                    items: { type: "string", format: "binary" },
                  },
                },
              },
            },
          },
        },
        responses: {
          "201": response("Report stored", {
            type: "object",
            properties: {
              message: { type: "string" },
              report_meta: { $ref: "#/components/schemas/ReportMetadata" },
              stats: { type: "object" },
              stored: { type: "object" },
            },
          }),
          "400": errorResponse("Invalid upload"),
          "401": errorResponse("Unauthorized"),
          "403": errorResponse("Administrator access required"),
          "422": errorResponse("Report transformation failed"),
        },
      },
    },
    "/api/reports": {
      get: {
        operationId: "listReportMonths",
        summary: "List available report months",
        responses: {
          "200": response("Available months", {
            type: "object",
            properties: {
              report_months: { type: "array", items: { type: "string", format: "date" } },
            },
          }),
          "401": errorResponse("Unauthorized"),
        },
      },
    },
    "/api/costs": {
      get: {
        operationId: "getCosts",
        summary: "Read filtered resolved cost rows",
        parameters: [reportMonthQuery, ...costFilters, limitQuery],
        responses: {
          "200": response("Cost rows", {
            type: "object",
            properties: {
              report_month: { type: "string", format: "date" },
              count: { type: "integer" },
              items: { type: "array", items: { $ref: "#/components/schemas/CostItem" } },
            },
          }),
          "400": errorResponse("Invalid query"),
          "401": errorResponse("Unauthorized"),
        },
      },
    },
    "/api/ai/usage": {
      get: {
        operationId: "getAiUsage",
        summary: "Read filtered AI usage",
        parameters: [reportMonthQuery, ...aiFilters, limitQuery],
        responses: {
          "200": response("AI usage rows", {
            type: "object",
            properties: {
              report_month: { type: "string", format: "date" },
              count: { type: "integer" },
              items: { type: "array", items: { $ref: "#/components/schemas/AiUsageItem" } },
            },
          }),
          "400": errorResponse("Invalid query"),
          "401": errorResponse("Unauthorized"),
        },
      },
    },
    "/api/overview/summary": {
      get: {
        operationId: "getOverviewSummary",
        summary: "Read filtered report KPIs",
        parameters: [reportMonthQuery, ...costFilters],
        responses: {
          "200": response("Report summary", {
            type: "object",
            properties: {
              report_month: { type: "string", format: "date" },
              billing_period_start: { type: "string", format: "date" },
              billing_period_end: { type: "string", format: "date" },
              currency: { type: "string" },
              total_cost: { type: "number" },
              subscription_count: { type: "integer" },
              resource_group_count: { type: "integer" },
              resource_count: { type: "integer" },
              line_item_count: { type: "integer" },
              highest_cost_item: { $ref: "#/components/schemas/CostItem" },
            },
          }),
          "400": errorResponse("Invalid query"),
          "401": errorResponse("Unauthorized"),
          "404": errorResponse("No matching report data"),
        },
      },
    },
    "/api/reports/{reportMonth}/excel": {
      get: {
        operationId: "downloadReportExcel",
        summary: "Generate a filtered Excel workbook from Supabase",
        parameters: [reportMonthPath, ...costFilters],
        responses: {
          "200": {
            description: "Generated XLSX workbook",
            content: {
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {
                schema: { type: "string", format: "binary" },
              },
            },
          },
          "400": errorResponse("Invalid query"),
          "401": errorResponse("Unauthorized"),
          "404": errorResponse("No matching report data"),
        },
      },
    },
    "/api/settings/project-applications": {
      get: {
        operationId: "listProjectApplications",
        summary: "List project/application pairs",
        responses: {
          "200": response("Project/application pairs", {
            type: "object",
            properties: {
              items: { type: "array", items: { $ref: "#/components/schemas/ProjectApplication" } },
            },
          }),
          "401": errorResponse("Unauthorized"),
        },
      },
      post: {
        operationId: "createProjectApplication",
        summary: "Create a project/application pair",
        description: "Admin only.",
        requestBody: { required: true, content: json(projectApplicationInput) },
        responses: {
          "201": response("Created", { $ref: "#/components/schemas/ProjectApplication" }),
          "400": errorResponse("Invalid input"),
          "401": errorResponse("Unauthorized"),
          "403": errorResponse("Administrator access required"),
          "409": errorResponse("Project/application pair already exists"),
        },
      },
    },
    "/api/settings/project-applications/{id}": {
      patch: {
        operationId: "updateProjectApplication",
        summary: "Update or deactivate a project/application pair",
        description: "Admin only. Deactivation also archives its active mappings.",
        parameters: [idPath],
        requestBody: {
          required: true,
          content: json({ ...projectApplicationInput, required: [], minProperties: 1 }),
        },
        responses: {
          "200": response("Updated", { $ref: "#/components/schemas/ProjectApplication" }),
          "400": errorResponse("Invalid input"),
          "401": errorResponse("Unauthorized"),
          "403": errorResponse("Administrator access required"),
          "409": errorResponse("Conflicting project/application pair"),
        },
      },
    },
    "/api/settings/resource-mappings": {
      get: {
        operationId: "listResourceMappings",
        summary: "List resource mappings",
        responses: {
          "200": response("Resource mappings", {
            type: "object",
            properties: {
              items: { type: "array", items: { $ref: "#/components/schemas/ResourceMapping" } },
            },
          }),
          "401": errorResponse("Unauthorized"),
        },
      },
      post: {
        operationId: "createResourceMapping",
        summary: "Create a resource-group mapping or resource override",
        description: "Admin only. Supply both resource_id and resource_name, or neither.",
        requestBody: { required: true, content: json(resourceMappingInput) },
        responses: {
          "201": response("Created", { $ref: "#/components/schemas/ResourceMapping" }),
          "400": errorResponse("Invalid input"),
          "401": errorResponse("Unauthorized"),
          "403": errorResponse("Administrator access required"),
          "409": errorResponse("Conflicting mapping or inactive application"),
        },
      },
    },
    "/api/settings/resource-mappings/{id}": {
      patch: {
        operationId: "updateResourceMapping",
        summary: "Update a resource mapping",
        description: "Admin only.",
        parameters: [idPath],
        requestBody: {
          required: true,
          content: json({ ...resourceMappingInput, required: [], minProperties: 1 }),
        },
        responses: {
          "200": response("Updated", { $ref: "#/components/schemas/ResourceMapping" }),
          "400": errorResponse("Invalid input"),
          "401": errorResponse("Unauthorized"),
          "403": errorResponse("Administrator access required"),
          "409": errorResponse("Conflicting mapping or inactive application"),
        },
      },
      delete: {
        operationId: "archiveResourceMapping",
        summary: "Archive a resource mapping",
        description: "Admin only. Sets is_active to false instead of hard-deleting the record.",
        parameters: [idPath],
        responses: {
          "200": response("Archived", { $ref: "#/components/schemas/ResourceMapping" }),
          "401": errorResponse("Unauthorized"),
          "403": errorResponse("Administrator access required"),
        },
      },
    },
    "/api/settings/resource-catalog": {
      get: {
        operationId: "getResourceCatalog",
        summary: "List discovered subscriptions, resource groups, and resources",
        parameters: [{ ...reportMonthQuery, required: false }],
        responses: {
          "200": response("Resource hierarchy", {
            type: "object",
            properties: { items: { type: "array", items: { type: "object" } } },
          }),
          "400": errorResponse("Invalid report month"),
          "401": errorResponse("Unauthorized"),
        },
      },
    },
    "/api/settings/unmapped": {
      get: {
        operationId: "getUnmappedResources",
        summary: "List cost rows without a configured mapping",
        parameters: [reportMonthQuery],
        responses: {
          "200": response("Unmapped cost rows", {
            type: "object",
            properties: { items: { type: "array", items: { type: "object" } } },
          }),
          "400": errorResponse("Invalid report month"),
          "401": errorResponse("Unauthorized"),
        },
      },
    },
  },
};
