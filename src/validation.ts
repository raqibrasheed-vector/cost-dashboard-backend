import { ApiError } from "./errors.js";

export function reportMonth(value: unknown): string {
  const text = String(value ?? "");
  if (!/^\d{4}-\d{2}-01$/.test(text) || Number.isNaN(Date.parse(`${text}T00:00:00Z`))) {
    throw new ApiError(400, "report_month must be the first day of a month in YYYY-MM-DD format.");
  }
  return text;
}

export function boundedLimit(value: unknown, fallback = 1000): number {
  if (value === undefined) return fallback;
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1 || parsed > 5000) {
    throw new ApiError(400, "limit must be an integer between 1 and 5000.");
  }
  return parsed;
}

export function requiredText(body: Record<string, unknown>, field: string): string {
  const value = String(body[field] ?? "").trim();
  if (!value) throw new ApiError(400, `${field} is required.`);
  return value;
}

export function optionalText(value: unknown): string | undefined {
  const text = String(value ?? "").trim();
  return text || undefined;
}

export function optionalBoolean(value: unknown, field: string): boolean | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "boolean") throw new ApiError(400, `${field} must be a boolean.`);
  return value;
}
