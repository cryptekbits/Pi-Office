import type { Request, Response, NextFunction } from "express";

const LOOPBACK_HOSTS = new Set(["localhost", "127.0.0.1", "::1"]);

export function isAllowedCompanionOrigin(origin: string | undefined): boolean {
  if (!origin) return false;
  try {
    const parsed = new URL(origin);
    return (parsed.protocol === "https:" || parsed.protocol === "http:") && LOOPBACK_HOSTS.has(parsed.hostname);
  } catch {
    return false;
  }
}

export function companionCorsMiddleware(request: Request, response: Response, next: NextFunction): void {
  const origin = request.header("origin");
  if (isAllowedCompanionOrigin(origin)) {
    response.setHeader("Access-Control-Allow-Origin", origin!);
    response.setHeader("Vary", "Origin");
    response.setHeader("Access-Control-Allow-Methods", "GET,POST,DELETE,OPTIONS");
    response.setHeader("Access-Control-Allow-Headers", "Content-Type");
  }

  if (request.method === "OPTIONS") {
    response.sendStatus(204);
    return;
  }

  next();
}
