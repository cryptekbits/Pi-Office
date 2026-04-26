import { randomUUID } from "node:crypto";
import { existsSync, mkdirSync } from "node:fs";
import { delimiter, isAbsolute, join, relative, resolve } from "node:path";
import { spawnSync } from "node:child_process";
import type { BashOperations } from "@mariozechner/pi-coding-agent";
import type {
  CompanionShellBackend,
  CompanionShellCapability,
  CompanionShellExecuteRequest,
  CompanionShellExecuteResponse,
  CompanionShellPolicy,
  CompanionShellProbeResult,
  CompanionShellState,
} from "@pi-office/pi-office-pack/protocol";

export const SHELL_POLICY_VERSION = "companion-shell-sandbox-v1";
const DEFAULT_TIMEOUT_MS = 10_000;
const DEFAULT_OUTPUT_BYTE_LIMIT = 50_000;

const DENIED_PATTERNS = [
  ".env",
  ".env.*",
  ".ssh",
  ".aws",
  ".azure",
  ".gcloud",
  ".kube",
  ".npmrc",
  ".pypirc",
  ".netrc",
  "id_rsa",
  "id_ed25519",
  "*.pem",
  "*.key",
  "*.pfx",
  "AppData/*/Microsoft/Office",
  "AppData/*/Google",
  "AppData/*/Mozilla",
];

const SECRET_PATTERN =
  /(^|[\\/])(?:\.env(?:\.|$)|\.ssh(?:[\\/]|$)|\.aws(?:[\\/]|$)|\.azure(?:[\\/]|$)|\.gcloud(?:[\\/]|$)|\.kube(?:[\\/]|$)|\.npmrc$|\.pypirc$|\.netrc$|id_rsa$|id_ed25519$)|\.(?:pem|key|pfx)\b|appdata[\\/].*(?:microsoft[\\/]office|google|mozilla)/i;
const NETWORK_PATTERN =
  /\b(?:curl|wget|iwr|invoke-webrequest|fetch|ssh|scp|sftp|rsync|nc|netcat|telnet|ping|nslookup|dig)\b|https?:\/\//i;
const PACKAGE_INSTALL_PATTERN =
  /\b(?:npm|pnpm|yarn|bun)\s+(?:i|install|add|dlx|exec)\b|\b(?:pip|pip3|uv)\s+(?:install|run)\b|\b(?:cargo|go)\s+(?:install|get)\b/i;
const VCS_MUTATION_PATTERN =
  /\bgit\s+(?:push|commit|reset|checkout|switch|rebase|merge|clean|tag|branch\s+-d|branch\s+-D)\b/i;
const MUTATING_COMMAND_PATTERN =
  /(^|[;&|]\s*)(?:rm|rmdir|del|erase|mv|move|cp|copy|touch|mkdir|chmod|chown|attrib|setx|reg|tee)\b|(^|[^<])>{1,2}(?!>)|(^|[;&|]\s*)git\s+(?:add|mv|rm)\b/i;

const SAFE_ENV_KEYS = new Set([
  "PATH",
  "Path",
  "SystemRoot",
  "WINDIR",
  "COMSPEC",
  "TMP",
  "TEMP",
  "TMPDIR",
  "LANG",
  "LC_ALL",
  "HOME",
  "USERPROFILE",
]);

export interface ShellBackendDetection {
  state: CompanionShellState;
  backend: CompanionShellBackend;
  reason: string;
}

export interface ShellSandboxRunnerRequest {
  command: string;
  cwd: string;
  timeoutMs: number;
  outputByteLimit: number;
  env: NodeJS.ProcessEnv;
  signal?: AbortSignal | undefined;
}

export interface ShellSandboxRunnerResult {
  exitCode: number | null;
  stdout: string;
  stderr: string;
  touchedPaths?: string[] | undefined;
  deniedPaths?: string[] | undefined;
  durationMs?: number | undefined;
  timedOut?: boolean | undefined;
}

export interface ShellSandboxRunner {
  backend: CompanionShellBackend;
  execute(request: ShellSandboxRunnerRequest): Promise<ShellSandboxRunnerResult>;
}

export interface CompanionShellSandboxOptions {
  sessionId: string;
  dataDir: string;
  workspaceDir?: string | undefined;
  detection?: ShellBackendDetection | undefined;
  runner?: ShellSandboxRunner | undefined;
  env?: NodeJS.ProcessEnv | undefined;
  now?: (() => Date) | undefined;
}

interface ValidationResult {
  ok: boolean;
  cwd: string;
  error?: string | undefined;
  deniedPaths: string[];
  touchedPaths: string[];
}

function defaultCommandExists(command: string): boolean {
  const executable = process.platform === "win32" && !command.endsWith(".exe") ? `${command}.exe` : command;
  const result = spawnSync(executable, ["--version"], {
    stdio: "ignore",
    shell: false,
    timeout: 1_500,
  });
  return !result.error && typeof result.status === "number";
}

export function detectShellSandboxBackend(options: {
  platform?: NodeJS.Platform | undefined;
  env?: NodeJS.ProcessEnv | undefined;
  commandExists?: ((command: string) => boolean) | undefined;
} = {}): ShellBackendDetection {
  const env = options.env ?? process.env;
  const platform = options.platform ?? process.platform;
  const commandExists = options.commandExists ?? defaultCommandExists;

  if (env.PI_OFFICE_ENABLE_SHELL_SANDBOX !== "1") {
    return {
      state: "unavailable",
      backend: "none",
      reason: "Companion shell is disabled by default. Set PI_OFFICE_ENABLE_SHELL_SANDBOX=1 only after choosing an isolation backend.",
    };
  }

  if (platform === "win32") {
    if (commandExists("docker")) {
      return {
        state: "degraded",
        backend: "docker",
        reason: "Docker is present, but Pi-Office does not enable shell until the container runner and destructive probes pass.",
      };
    }
    if (commandExists("wsl") || commandExists("wsl.exe")) {
      return {
        state: "degraded",
        backend: "wsl2",
        reason: "WSL is present, but native Windows-to-WSL path and network isolation probes have not passed.",
      };
    }
    return {
      state: "unavailable",
      backend: "none",
      reason: "No supported Windows isolation backend was detected. Raw PowerShell, cmd, and host bash stay unavailable.",
    };
  }

  if (platform === "linux") {
    if (commandExists("bwrap")) {
      return {
        state: "degraded",
        backend: "linux-bubblewrap",
        reason: "bubblewrap is present, but Pi-Office requires a runner-specific probe pass before exposing shell.",
      };
    }
    if (commandExists("docker")) {
      return {
        state: "degraded",
        backend: "docker",
        reason: "Docker is present, but Pi-Office does not enable shell until the container runner and destructive probes pass.",
      };
    }
    return {
      state: "unavailable",
      backend: "none",
      reason: "No supported Linux isolation backend was detected. Raw host bash stays unavailable.",
    };
  }

  if (platform === "darwin") {
    if (commandExists("docker")) {
      return {
        state: "degraded",
        backend: "docker",
        reason: "Docker is present, but Pi-Office does not enable shell until the container runner and destructive probes pass.",
      };
    }
    return {
      state: "unavailable",
      backend: "none",
      reason: "No supported macOS isolation backend was detected. Native shell remains unavailable until proven safe.",
    };
  }

  return {
    state: "unavailable",
    backend: "none",
    reason: `Unsupported platform for companion shell sandbox: ${platform}.`,
  };
}

function normalizeComparablePath(value: string): string {
  const resolved = resolve(value);
  return process.platform === "win32" ? resolved.toLowerCase() : resolved;
}

function isPathInside(root: string, candidate: string): boolean {
  const resolvedRoot = normalizeComparablePath(root);
  const resolvedCandidate = normalizeComparablePath(candidate);
  const relativePath = relative(resolvedRoot, resolvedCandidate);
  return relativePath === "" || (!relativePath.startsWith("..") && !isAbsolute(relativePath));
}

function hasDeniedSecretReference(value: string): boolean {
  return SECRET_PATTERN.test(value.replaceAll("\\", "/"));
}

function capText(value: string, maxBytes: number): { text: string; capped: boolean } {
  const buffer = Buffer.from(value, "utf8");
  if (buffer.byteLength <= maxBytes) {
    return { text: value, capped: false };
  }
  return {
    text: `${buffer.subarray(0, maxBytes).toString("utf8")}\n[Pi-Office shell output capped at ${maxBytes} bytes]`,
    capped: true,
  };
}

function scrubEnvironment(env: NodeJS.ProcessEnv): NodeJS.ProcessEnv {
  const next: NodeJS.ProcessEnv = {};
  for (const [key, value] of Object.entries(env)) {
    if (value !== undefined && SAFE_ENV_KEYS.has(key)) {
      next[key] = value;
    }
  }
  return next;
}

function extractPathCandidates(command: string): string[] {
  const candidates = new Set<string>();
  const quotedPattern = /["']([^"']*[\\/][^"']*|\.env(?:\.[^"']*)?|\.ssh[^"']*)["']/g;
  const tokenPattern = /(?:^|\s)(\.{1,2}[\\/][^\s;&|]+|[A-Za-z]:[\\/][^\s;&|]+|\/[^\s;&|]+|\.env(?:\.[^\s;&|]+)?|\.ssh(?:[\\/][^\s;&|]+)?)/g;
  let match: RegExpExecArray | null;
  while ((match = quotedPattern.exec(command)) !== null) {
    if (match[1]) candidates.add(match[1]);
  }
  while ((match = tokenPattern.exec(command)) !== null) {
    if (match[1]) candidates.add(match[1]);
  }
  return Array.from(candidates);
}

function resolveCommandCwd(policy: CompanionShellPolicy, request: CompanionShellExecuteRequest): string {
  const fallback = request.category === "scratch-write"
    ? policy.scratchRoot
    : policy.readableRoots[0] ?? policy.scratchRoot;
  if (!fallback) {
    throw new Error("Shell requires a saved document folder or scratch root.");
  }
  return resolve(request.cwd ?? fallback);
}

function classifyCommand(policy: CompanionShellPolicy, request: CompanionShellExecuteRequest): string | undefined {
  const command = request.command.trim();
  if (!command) return "Shell command is required.";
  if (hasDeniedSecretReference(command)) return "Command references denied secret paths or credential-like files.";
  if (policy.network === "disabled" && NETWORK_PATTERN.test(command)) return "Network commands are disabled by the companion shell policy.";
  if (PACKAGE_INSTALL_PATTERN.test(command)) return "Package installation and package-manager execution are disabled by the companion shell policy.";
  if (VCS_MUTATION_PATTERN.test(command)) return "Mutating git commands are disabled by the companion shell policy.";
  if (MUTATING_COMMAND_PATTERN.test(command) && request.category !== "scratch-write") {
    return "Mutating shell commands require the scratch-write category and scratch working directory.";
  }
  return undefined;
}

export function createShellPolicy(options: {
  sessionId: string;
  dataDir: string;
  workspaceDir?: string | undefined;
}): CompanionShellPolicy {
  const readableRoots = options.workspaceDir ? [resolve(options.workspaceDir)] : [];
  return {
    version: SHELL_POLICY_VERSION,
    readableRoots,
    scratchRoot: resolve(join(options.dataDir, "shell-scratch", options.sessionId)),
    deniedPatterns: [...DENIED_PATTERNS],
    network: "disabled",
    timeoutMs: DEFAULT_TIMEOUT_MS,
    outputByteLimit: DEFAULT_OUTPUT_BYTE_LIMIT,
  };
}

export function validateShellRequest(policy: CompanionShellPolicy, request: CompanionShellExecuteRequest): ValidationResult {
  const deniedPaths: string[] = [];
  const touchedPaths: string[] = [];
  let cwd: string;
  try {
    cwd = resolveCommandCwd(policy, request);
  } catch (error) {
    return {
      ok: false,
      cwd: ".",
      error: error instanceof Error ? error.message : String(error),
      deniedPaths,
      touchedPaths,
    };
  }

  const readableOrScratchRoots = [
    ...policy.readableRoots,
    ...(policy.scratchRoot ? [policy.scratchRoot] : []),
  ];
  const cwdAllowed = request.category === "scratch-write"
    ? Boolean(policy.scratchRoot && isPathInside(policy.scratchRoot, cwd))
    : readableOrScratchRoots.some((root) => isPathInside(root, cwd));
  if (!cwdAllowed) {
    deniedPaths.push(cwd);
    return {
      ok: false,
      cwd,
      error: request.category === "scratch-write"
        ? "Scratch-write commands must run inside the companion scratch directory."
        : "Shell working directory must stay inside a readable root or scratch directory.",
      deniedPaths,
      touchedPaths,
    };
  }

  const commandError = classifyCommand(policy, request);
  if (commandError) {
    return { ok: false, cwd, error: commandError, deniedPaths, touchedPaths };
  }

  for (const candidate of extractPathCandidates(request.command)) {
    if (hasDeniedSecretReference(candidate)) {
      deniedPaths.push(candidate);
      return {
        ok: false,
        cwd,
        error: "Command references denied secret paths or credential-like files.",
        deniedPaths,
        touchedPaths,
      };
    }

    const resolvedCandidate = resolve(cwd, candidate);
    const allowed = request.category === "scratch-write"
      ? Boolean(policy.scratchRoot && isPathInside(policy.scratchRoot, resolvedCandidate))
      : readableOrScratchRoots.some((root) => isPathInside(root, resolvedCandidate));
    if (!allowed) {
      deniedPaths.push(resolvedCandidate);
      return {
        ok: false,
        cwd,
        error: "Command path references must stay inside readable roots or scratch.",
        deniedPaths,
        touchedPaths,
      };
    }
    touchedPaths.push(resolvedCandidate);
  }

  return { ok: true, cwd, deniedPaths, touchedPaths };
}

export function runShellPolicyProbes(policy: CompanionShellPolicy): CompanionShellProbeResult[] {
  const workspaceRoot = policy.readableRoots[0] ?? resolve(policy.scratchRoot ?? ".");
  const scratchRoot = policy.scratchRoot ?? resolve(workspaceRoot, ".pi-office-shell-scratch");
  const outsidePath = resolve(workspaceRoot, "..", "outside-pi-office-probe.txt");
  const probes: Array<{ id: string; description: string; request: CompanionShellExecuteRequest }> = [
    {
      id: "deny-write-outside-scratch",
      description: "Writing outside scratch is denied before execution.",
      request: { command: `echo bad > "${outsidePath}"`, cwd: workspaceRoot, category: "read-only" },
    },
    {
      id: "deny-delete-outside-scratch",
      description: "Deleting outside scratch is denied before execution.",
      request: { command: `rm "${outsidePath}"`, cwd: workspaceRoot, category: "read-only" },
    },
    {
      id: "deny-env-read",
      description: "Reading .env files is denied before execution.",
      request: { command: "cat .env", cwd: workspaceRoot, category: "read-only" },
    },
    {
      id: "deny-symlink-secret-escape",
      description: "Symlink or junction escape probes targeting secrets are denied before execution.",
      request: { command: `cat "${scratchRoot}/link-to-secret/.env"`, cwd: scratchRoot, category: "read-only" },
    },
    {
      id: "deny-network",
      description: "Network clients are denied before execution.",
      request: { command: "curl https://example.com", cwd: workspaceRoot, category: "read-only" },
    },
    {
      id: "deny-package-install",
      description: "Package installs are denied before execution.",
      request: { command: "npm install left-pad", cwd: workspaceRoot, category: "read-only" },
    },
    {
      id: "deny-git-push",
      description: "git push is denied before execution.",
      request: { command: "git push origin main", cwd: workspaceRoot, category: "read-only" },
    },
    {
      id: "deny-git-commit",
      description: "git commit is denied before execution.",
      request: { command: "git commit -am probe", cwd: workspaceRoot, category: "read-only" },
    },
  ];

  return probes.map((probe) => {
    const result = validateShellRequest(policy, probe.request);
    return {
      id: probe.id,
      description: probe.description,
      ok: !result.ok,
      blocked: !result.ok,
      detail: result.error ?? "Probe was not blocked.",
    };
  });
}

export class CompanionShellSandbox {
  readonly policy: CompanionShellPolicy;
  private readonly detection: ShellBackendDetection;
  private readonly runner: ShellSandboxRunner | undefined;
  private readonly now: () => Date;
  private readonly env: NodeJS.ProcessEnv;
  private capability: CompanionShellCapability | undefined;

  constructor(options: CompanionShellSandboxOptions) {
    this.policy = createShellPolicy({
      sessionId: options.sessionId,
      dataDir: options.dataDir,
      workspaceDir: options.workspaceDir,
    });
    this.detection = options.detection ?? detectShellSandboxBackend({ env: options.env });
    this.runner = options.runner;
    this.now = options.now ?? (() => new Date());
    this.env = options.env ?? process.env;
  }

  getCapability(): CompanionShellCapability {
    if (this.capability) return this.capability;

    const policyProbes = runShellPolicyProbes(this.policy);
    let state = this.detection.state;
    let reason = this.detection.reason;
    let backend = this.detection.backend;

    if (!this.runner) {
      state = this.detection.state === "degraded" ? "degraded" : "unavailable";
      reason = this.detection.reason;
    } else if (this.detection.state === "unavailable") {
      state = "unavailable";
      reason = this.detection.reason;
    } else if (policyProbes.every((probe) => probe.ok)) {
      state = "available";
      backend = this.runner.backend;
      reason = "Companion shell sandbox runner is available and policy probes passed.";
    } else {
      state = "degraded";
      reason = "Companion shell sandbox policy probes did not pass; shell execution is disabled.";
    }

    this.capability = {
      state,
      backend,
      detectedAt: this.now().toISOString(),
      reason,
      policy: this.policy,
      probes: policyProbes,
    };
    return this.capability;
  }

  async execute(request: CompanionShellExecuteRequest, signal?: AbortSignal | undefined): Promise<CompanionShellExecuteResponse> {
    const startedAt = Date.now();
    const auditId = randomUUID();
    const capability = this.getCapability();
    const validation = validateShellRequest(this.policy, request);
    const duration = () => Date.now() - startedAt;

    const base = {
      auditId,
      backend: capability.backend,
      policyVersion: this.policy.version,
      stdout: "",
      stderr: "",
      capped: false,
      deniedPaths: validation.deniedPaths,
      touchedPaths: validation.touchedPaths,
    };

    if (!validation.ok) {
      return {
        ...base,
        ok: false,
        durationMs: duration(),
        exitCode: 126,
        error: validation.error ?? "Shell command denied by policy.",
        stderr: validation.error ?? "Shell command denied by policy.",
      };
    }

    if (capability.state !== "available" || !this.runner) {
      const message = capability.reason ?? "Companion shell sandbox is unavailable.";
      return {
        ...base,
        ok: false,
        durationMs: duration(),
        exitCode: 126,
        error: message,
        stderr: message,
      };
    }

    if (this.policy.scratchRoot) {
      mkdirSync(this.policy.scratchRoot, { recursive: true });
    }

    const timeoutMs = Math.min(
      Math.max(request.timeoutMs ?? this.policy.timeoutMs, 1_000),
      this.policy.timeoutMs,
    );

    const result = await this.runner.execute({
      command: request.command.trim(),
      cwd: validation.cwd,
      timeoutMs,
      outputByteLimit: this.policy.outputByteLimit,
      env: scrubEnvironment(this.env),
      signal,
    });
    const stdout = capText(result.stdout, this.policy.outputByteLimit);
    const stderr = capText(result.stderr, this.policy.outputByteLimit);
    const timedOutMessage = result.timedOut ? `Command timed out after ${timeoutMs}ms.` : undefined;

    return {
      ...base,
      ok: result.exitCode === 0 && !result.timedOut,
      durationMs: result.durationMs ?? duration(),
      exitCode: result.exitCode,
      stdout: stdout.text,
      stderr: timedOutMessage ? [stderr.text, timedOutMessage].filter(Boolean).join("\n") : stderr.text,
      capped: stdout.capped || stderr.capped,
      deniedPaths: [...validation.deniedPaths, ...(result.deniedPaths ?? [])],
      touchedPaths: [...validation.touchedPaths, ...(result.touchedPaths ?? [])],
      ...(timedOutMessage ? { error: timedOutMessage } : {}),
    };
  }
}

export async function runShellExecutionProbes(sandbox: CompanionShellSandbox): Promise<CompanionShellProbeResult[]> {
  const scratchRoot = sandbox.policy.scratchRoot ?? ".";
  const timeoutProbe = await sandbox.execute({
    command: "pi-office-probe-timeout",
    cwd: scratchRoot,
    category: "scratch-write",
    timeoutMs: 1_000,
  });
  const outputProbe = await sandbox.execute({
    command: "pi-office-probe-output-cap",
    cwd: scratchRoot,
    category: "scratch-write",
    timeoutMs: 1_000,
  });

  return [
    {
      id: "enforce-timeout",
      description: "Long-running commands terminate on timeout.",
      ok: !timeoutProbe.ok && /timed out/i.test(timeoutProbe.error ?? timeoutProbe.stderr),
      blocked: !timeoutProbe.ok,
      detail: timeoutProbe.error ?? (timeoutProbe.stderr || "Timeout probe did not fail."),
    },
    {
      id: "enforce-output-cap",
      description: "Noisy commands are capped before returning to the model.",
      ok: outputProbe.capped,
      blocked: false,
      detail: outputProbe.capped ? "Output was capped." : "Output was not capped.",
    },
  ];
}

export function createCompanionBashOperations(sandbox: CompanionShellSandbox): BashOperations {
  return {
    async exec(command, cwd, options) {
      const response = await sandbox.execute(
        {
          command,
          cwd,
          category: "read-only",
          timeoutMs: options.timeout ? options.timeout * 1_000 : undefined,
        },
        options.signal,
      );
      const output = [response.stdout, response.stderr, response.error]
        .filter((value): value is string => Boolean(value))
        .join("\n");
      if (output) {
        options.onData(Buffer.from(output));
      }
      if (options.signal?.aborted) {
        throw new Error("aborted");
      }
      return {
        exitCode: response.exitCode ?? (response.ok ? 0 : 1),
      };
    },
  };
}

export function pathListForDiagnostics(policy: CompanionShellPolicy): string {
  return [
    ...policy.readableRoots,
    ...(policy.scratchRoot ? [policy.scratchRoot] : []),
  ].join(delimiter);
}

export function shellSandboxExistsForSession(sandbox: CompanionShellSandbox): boolean {
  const scratchRoot = sandbox.policy.scratchRoot;
  return Boolean(scratchRoot && existsSync(resolve(scratchRoot)));
}
