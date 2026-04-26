import { spawnSync } from "node:child_process";
import type {
  ConnectorDiagnosticsResponse,
  ConnectorRuntimeCheck,
  ConnectorRuntimeCheckKey,
} from "@pi-office/pi-office-pack/protocol";

interface RuntimeProbe {
  key: ConnectorRuntimeCheckKey;
  label: string;
  commands: string[];
}

export interface CommandProbeResult {
  ok: boolean;
  detail: string;
}

export type CommandProbeRunner = (command: string) => CommandProbeResult;

const RUNTIME_PROBES: RuntimeProbe[] = [
  { key: "node", label: "Node.js", commands: ["node"] },
  { key: "npm", label: "npm", commands: ["npm"] },
  { key: "npx", label: "npx", commands: ["npx"] },
  { key: "python", label: "Python", commands: ["python", "python3"] },
  { key: "uv", label: "uv", commands: ["uv"] },
  { key: "docker", label: "Docker", commands: ["docker"] },
  { key: "git", label: "Git", commands: ["git"] },
];

function commandCandidates(command: string, platform = process.platform): string[] {
  if (platform !== "win32") return [command];
  if (/\.(?:exe|cmd|bat)$/i.test(command)) return [command];
  return [command, `${command}.cmd`, `${command}.exe`, `${command}.bat`];
}

export function probeCommandVersion(command: string, platform = process.platform): CommandProbeResult {
  let lastError = "not found";
  for (const candidate of commandCandidates(command, platform)) {
    const result = spawnSync(candidate, ["--version"], {
      encoding: "utf8",
      shell: false,
      timeout: 2_500,
    });
    const output = `${result.stdout ?? ""}${result.stderr ?? ""}`.trim().split(/\r?\n/)[0]?.trim();
    if (!result.error && result.status === 0) {
      return {
        ok: true,
        detail: output ? `${candidate}: ${output}` : `${candidate} is available.`,
      };
    }
    if (result.error?.message) {
      lastError = result.error.message;
    } else if (typeof result.status === "number") {
      lastError = `${candidate} exited with ${result.status}`;
    }
  }
  return { ok: false, detail: lastError };
}

export function createCompanionRuntimeDiagnostics(
  runner: CommandProbeRunner = (command) => probeCommandVersion(command),
  now: () => Date = () => new Date(),
): ConnectorDiagnosticsResponse {
  const runtimes: ConnectorRuntimeCheck[] = RUNTIME_PROBES.map((probe) => {
    for (const command of probe.commands) {
      const result = runner(command);
      if (result.ok) {
        return {
          key: probe.key,
          label: probe.label,
          ok: true,
          detail: result.detail,
        };
      }
    }
    const fallback = runner(probe.commands[0]!);
    return {
      key: probe.key,
      label: probe.label,
      ok: false,
      detail: fallback.detail,
    };
  });

  return {
    generatedAt: now().toISOString(),
    runtimes,
    envSuggestions: [],
    diagnostics: [
      {
        level: "info",
        code: "companion_runtime",
        title: "Companion runtime",
        message: "Runtime checks were collected from the optional local companion process.",
      },
    ],
  };
}
