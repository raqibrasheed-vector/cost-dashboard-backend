import dotenv from "dotenv";
import path from "node:path";
import { fileURLToPath } from "node:url";

dotenv.config({ quiet: true });

const apiRoot = path.resolve(fileURLToPath(new URL("..", import.meta.url)));

function positiveInteger(value: string | undefined, fallback: number): number {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

export const config = {
  apiRoot,
  port: positiveInteger(process.env.PORT, 3001),
  corsOrigin: process.env.CORS_ORIGIN?.trim() || "http://localhost:5173",
  uploadRoot: path.resolve(apiRoot, ".tmp", "uploads"),
};

export function requiredEnvironment(name: string): string {
  const value = process.env[name]?.trim();
  if (!value) {
    throw new Error(`${name} is not configured.`);
  }
  return value;
}
