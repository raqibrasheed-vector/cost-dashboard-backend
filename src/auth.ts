import type { NextFunction, Request, Response } from "express";

import { ApiError } from "./errors.js";
import { getSupabase } from "./database.js";

export async function requireUser(
  request: Request,
  response: Response,
  next: NextFunction,
) {
  try {
    const authorization = request.header("authorization") ?? "";
    const token = authorization.startsWith("Bearer ") ? authorization.slice(7) : "";
    if (!token) throw new ApiError(401, "A Supabase access token is required.");

    const { data, error } = await getSupabase().auth.getUser(token);
    if (error || !data.user) throw new ApiError(401, "Invalid Supabase access token.");
    response.locals.user = data.user;
    next();
  } catch (error) {
    next(error);
  }
}

export function requireAdmin(
  _request: Request,
  response: Response,
  next: NextFunction,
) {
  const user = response.locals.user;
  if (!user) {
    next(new ApiError(401, "A Supabase access token is required."));
    return;
  }
  if (user.app_metadata?.role !== "admin") {
    next(new ApiError(403, "Administrator access is required."));
    return;
  }
  next();
}
