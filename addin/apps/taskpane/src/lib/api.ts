import { dispatchKernelRequest } from "./runtime/inprocess-kernel";

export type RuntimeDiagnosticSource = "route" | "auth" | "connector";

export interface RuntimeRequestFailureDiagnostic {
  id: string;
  source: RuntimeDiagnosticSource;
  path: string;
  method: string;
  message: string;
  timestamp: number;
}

export const RUNTIME_DIAGNOSTIC_EVENT = "pi-runtime-diagnostic";

function resolveDiagnosticSource(path: string): RuntimeDiagnosticSource {
  if (path.startsWith("/v1/auth")) return "auth";
  if (path.startsWith("/v1/connectors")) return "connector";
  return "route";
}

function emitRuntimeDiagnostic(detail: RuntimeRequestFailureDiagnostic): void {
  if (typeof window === "undefined" || typeof window.dispatchEvent !== "function") return;
  window.dispatchEvent(new CustomEvent<RuntimeRequestFailureDiagnostic>(RUNTIME_DIAGNOSTIC_EVENT, { detail }));
}

export async function fetchJson<T>(path: string, init?: RequestInit): Promise<T> {
  try {
    return await dispatchKernelRequest<T>(path, init);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    emitRuntimeDiagnostic({
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      source: resolveDiagnosticSource(path),
      path,
      method: (init?.method ?? "GET").toUpperCase(),
      message,
      timestamp: Date.now(),
    });
    throw error;
  }
}

export async function postJson<T>(path: string, body: unknown): Promise<T> {
  return fetchJson<T>(path, {
    method: "POST",
    body: JSON.stringify(body),
  });
}

export async function deleteJson<T>(path: string): Promise<T> {
  return fetchJson<T>(path, {
    method: "DELETE",
  });
}
