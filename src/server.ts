import cors from "cors";
import express from "express";
import { randomUUID } from "node:crypto";
import { mkdir, rm } from "node:fs/promises";
import path from "node:path";
import multer from "multer";
import swaggerUi from "swagger-ui-express";

import { requireAdmin, requireUser } from "./auth.js";
import { config } from "./config.js";
import { ApiError, asyncRoute, errorHandler } from "./errors.js";
import {
  archiveResourceMapping,
  createProjectApplication,
  createResourceMapping,
  fetchAiUsage,
  fetchCostItems,
  listProjectApplications,
  listReportMonths,
  listResourceMappings,
  reportMetadata,
  resourceCatalog,
  storeTransformedReport,
  summarizeCosts,
  unmappedResources,
  updateProjectApplication,
  updateResourceMapping,
} from "./database.js";
import { openApiDocument } from "./openapi.js";
import { buildReportWorkbook } from "./report-workbook.js";
import { transformReportUploads } from "./report-transformer.js";
import type { AiUsageFilters, CostFilters } from "./types.js";
import {
  boundedLimit,
  optionalBoolean,
  optionalText,
  reportMonth,
  requiredText,
} from "./validation.js";

const excelMediaType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

function costFilters(query: Record<string, unknown>): CostFilters {
  return {
    subscription: optionalText(query.subscription),
    resource_group: optionalText(query.resource_group),
    project: optionalText(query.project),
    application: optionalText(query.application),
    environment: optionalText(query.environment),
    resource_type: optionalText(query.resource_type),
  };
}

function aiFilters(query: Record<string, unknown>): AiUsageFilters {
  return {
    subscription: optionalText(query.subscription),
    resource_group: optionalText(query.resource_group),
    project: optionalText(query.project),
    application: optionalText(query.application),
    environment: optionalText(query.environment),
    model_type: optionalText(query.model_type),
  };
}

function reportFilename(start: string, end: string): string {
  const startDate = new Date(`${start}T00:00:00Z`);
  const endDate = new Date(`${end}T00:00:00Z`);
  const month = startDate.toLocaleString("en-US", { month: "long", timeZone: "UTC" });
  return `Azure Cost Report ${month} ${startDate.getUTCDate()} - ${endDate.getUTCDate()}, ${endDate.getUTCFullYear()}.xlsx`;
}

function projectApplicationInput(body: Record<string, unknown>) {
  return {
    project_name: requiredText(body, "project_name"),
    application_name: requiredText(body, "application_name"),
    description: optionalText(body.description) ?? null,
    is_active: optionalBoolean(body.is_active, "is_active") ?? true,
  };
}

function projectApplicationPatch(body: Record<string, unknown>) {
  const patch: Record<string, unknown> = {};
  if (body.project_name !== undefined) patch.project_name = requiredText(body, "project_name");
  if (body.application_name !== undefined) patch.application_name = requiredText(body, "application_name");
  if (body.description !== undefined) patch.description = optionalText(body.description) ?? null;
  if (body.is_active !== undefined) patch.is_active = optionalBoolean(body.is_active, "is_active");
  if (!Object.keys(patch).length) throw new ApiError(400, "No supported fields were supplied.");
  return patch;
}

function mappingInput(body: Record<string, unknown>) {
  const resourceId = optionalText(body.resource_id) ?? null;
  const resourceName = optionalText(body.resource_name) ?? null;
  if ((resourceId === null) !== (resourceName === null)) {
    throw new ApiError(400, "resource_id and resource_name must be supplied together.");
  }
  return {
    project_application_id: requiredText(body, "project_application_id"),
    subscription_id: requiredText(body, "subscription_id"),
    subscription_name: optionalText(body.subscription_name) ?? null,
    resource_group_id: requiredText(body, "resource_group_id"),
    resource_group_name: optionalText(body.resource_group_name) ?? null,
    resource_id: resourceId,
    resource_name: resourceName,
    environment: requiredText(body, "environment"),
    is_active: optionalBoolean(body.is_active, "is_active") ?? true,
  };
}

function mappingPatch(body: Record<string, unknown>) {
  const patch: Record<string, unknown> = {};
  const requiredFields = [
    "project_application_id",
    "subscription_id",
    "resource_group_id",
    "environment",
  ];
  for (const field of requiredFields) {
    if (body[field] !== undefined) patch[field] = requiredText(body, field);
  }
  for (const field of ["subscription_name", "resource_group_name"]) {
    if (body[field] !== undefined) patch[field] = optionalText(body[field]) ?? null;
  }
  const hasResourceId = body.resource_id !== undefined;
  const hasResourceName = body.resource_name !== undefined;
  if (hasResourceId !== hasResourceName) {
    throw new ApiError(400, "resource_id and resource_name must be supplied together.");
  }
  if (hasResourceId) {
    const resourceId = optionalText(body.resource_id) ?? null;
    const resourceName = optionalText(body.resource_name) ?? null;
    if ((resourceId === null) !== (resourceName === null)) {
      throw new ApiError(400, "resource_id and resource_name must both be values or both be null.");
    }
    patch.resource_id = resourceId;
    patch.resource_name = resourceName;
  }
  if (body.is_active !== undefined) patch.is_active = optionalBoolean(body.is_active, "is_active");
  if (!Object.keys(patch).length) throw new ApiError(400, "No supported fields were supplied.");
  return patch;
}

export async function createApp() {
  await mkdir(config.uploadRoot, { recursive: true });
  const upload = multer({
    storage: multer.diskStorage({
      destination: config.uploadRoot,
      filename: (_request, file, callback) => {
        callback(null, `${randomUUID()}${path.extname(file.originalname).toLowerCase()}`);
      },
    }),
    limits: { fileSize: 50 * 1024 * 1024, files: 40 },
  });
  const app = express();
  app.disable("x-powered-by");
  app.use(cors({ origin: config.corsOrigin, credentials: true }));
  app.use(express.json({ limit: "1mb" }));

  app.get("/health", (_request, response) => {
    response.json({ status: "ok", runtime: "node", report_engine: "typescript" });
  });
  app.get("/openapi.json", (_request, response) => response.json(openApiDocument));
  app.use("/docs", swaggerUi.serve, swaggerUi.setup(openApiDocument));
  app.use("/api", requireUser);

  app.post(
    "/api/reports/ingest",
    requireAdmin,
    upload.fields([
      { name: "cost_reports", maxCount: 20 },
      { name: "ai_usage", maxCount: 20 },
    ]),
    asyncRoute(async (request, response) => {
      const files = (request.files ?? {}) as Record<string, Express.Multer.File[]>;
      const costFiles = files.cost_reports ?? [];
      const aiFiles = files.ai_usage ?? [];
      const allFiles = [...costFiles, ...aiFiles];
      try {
        if (!costFiles.length) throw new ApiError(400, "At least one cost_reports file is required.");
        if (costFiles.some((file) => path.extname(file.originalname).toLowerCase() !== ".xlsx")) {
          throw new ApiError(400, "Every cost report must be an .xlsx file.");
        }
        if (aiFiles.some((file) => path.extname(file.originalname).toLowerCase() !== ".csv")) {
          throw new ApiError(400, "Every AI usage file must be a .csv file.");
        }
        let transformed;
        try {
          transformed = await transformReportUploads(
            costFiles.map((file) => file.path),
            aiFiles.map((file) => file.path),
          );
        } catch (error) {
          const message = error instanceof Error ? error.message : "The report files are invalid.";
          throw new ApiError(422, message);
        }
        const stored = await storeTransformedReport(transformed);
        response.status(201).json({
          message: "Report data ingested successfully.",
          report_meta: transformed.report_meta,
          stats: transformed.stats,
          stored,
        });
      } finally {
        await Promise.all(allFiles.map((file) => rm(file.path, { force: true })));
      }
    }),
  );

  app.get(
    "/api/reports",
    asyncRoute(async (_request, response) => {
      response.json({ report_months: await listReportMonths() });
    }),
  );

  app.get(
    "/api/costs",
    asyncRoute(async (request, response) => {
      const month = reportMonth(request.query.report_month);
      const items = await fetchCostItems(month, costFilters(request.query), boundedLimit(request.query.limit));
      response.json({ report_month: month, count: items.length, items });
    }),
  );

  app.get(
    "/api/ai/usage",
    asyncRoute(async (request, response) => {
      const month = reportMonth(request.query.report_month);
      const items = await fetchAiUsage(month, aiFilters(request.query), boundedLimit(request.query.limit));
      response.json({ report_month: month, count: items.length, items });
    }),
  );

  app.get(
    "/api/overview/summary",
    asyncRoute(async (request, response) => {
      const month = reportMonth(request.query.report_month);
      response.json(summarizeCosts(await fetchCostItems(month, costFilters(request.query))));
    }),
  );

  app.get(
    "/api/reports/:reportMonth/excel",
    asyncRoute(async (request, response) => {
      const month = reportMonth(request.params.reportMonth);
      const filters = costFilters(request.query);
      const costItems = await fetchCostItems(month, filters);
      const aiUsage = await fetchAiUsage(month, {
        subscription: filters.subscription,
        resource_group: filters.resource_group,
        project: filters.project,
        application: filters.application,
        environment: filters.environment,
      });
      const metadata = reportMetadata(costItems);
      const workbook = await buildReportWorkbook(metadata, costItems, aiUsage);
      response
        .status(200)
        .type(excelMediaType)
        .attachment(reportFilename(metadata.period_start, metadata.period_end))
        .send(workbook);
    }),
  );

  app.get(
    "/api/settings/project-applications",
    asyncRoute(async (_request, response) => response.json({ items: await listProjectApplications() })),
  );
  app.post(
    "/api/settings/project-applications",
    requireAdmin,
    asyncRoute(async (request, response) => {
      response.status(201).json(await createProjectApplication(projectApplicationInput(request.body)));
    }),
  );
  app.patch(
    "/api/settings/project-applications/:id",
    requireAdmin,
    asyncRoute(async (request, response) => {
      response.json(
        await updateProjectApplication(String(request.params.id), projectApplicationPatch(request.body)),
      );
    }),
  );

  app.get(
    "/api/settings/resource-mappings",
    asyncRoute(async (_request, response) => response.json({ items: await listResourceMappings() })),
  );
  app.post(
    "/api/settings/resource-mappings",
    requireAdmin,
    asyncRoute(async (request, response) => {
      response.status(201).json(await createResourceMapping(mappingInput(request.body)));
    }),
  );
  app.patch(
    "/api/settings/resource-mappings/:id",
    requireAdmin,
    asyncRoute(async (request, response) => {
      response.json(
        await updateResourceMapping(String(request.params.id), mappingPatch(request.body)),
      );
    }),
  );
  app.delete(
    "/api/settings/resource-mappings/:id",
    requireAdmin,
    asyncRoute(async (request, response) => {
      response.json(await archiveResourceMapping(String(request.params.id)));
    }),
  );

  app.get(
    "/api/settings/resource-catalog",
    asyncRoute(async (request, response) => {
      const month = request.query.report_month ? reportMonth(request.query.report_month) : undefined;
      response.json({ items: await resourceCatalog(month) });
    }),
  );
  app.get(
    "/api/settings/unmapped",
    asyncRoute(async (request, response) => {
      response.json({ items: await unmappedResources(reportMonth(request.query.report_month)) });
    }),
  );

  app.use((_request, _response, next) => next(new ApiError(404, "Route not found.")));
  app.use(errorHandler);
  return app;
}

if (process.env.NODE_ENV !== "test") {
  const app = await createApp();
  app.listen(config.port, "127.0.0.1", () => {
    console.log(`Cost Report API listening on http://127.0.0.1:${config.port}`);
  });
}
