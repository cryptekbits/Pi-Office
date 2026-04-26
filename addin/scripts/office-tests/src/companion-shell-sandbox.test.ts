import assert from "node:assert/strict";
import test from "node:test";
import { resolve } from "node:path";
import {
  CompanionShellSandbox,
  createCompanionBashOperations,
  createShellPolicy,
  detectShellSandboxBackend,
  runShellExecutionProbes,
  runShellPolicyProbes,
  validateShellRequest,
  type ShellSandboxRunner,
  type ShellSandboxRunnerRequest,
  type ShellSandboxRunnerResult,
} from "../../../../companion/src/shell-sandbox.js";

class FakeShellRunner implements ShellSandboxRunner {
  readonly backend = "test" as const;
  readonly calls: ShellSandboxRunnerRequest[] = [];

  async execute(request: ShellSandboxRunnerRequest): Promise<ShellSandboxRunnerResult> {
    this.calls.push(request);
    if (request.command === "pi-office-probe-timeout") {
      return {
        exitCode: null,
        stdout: "",
        stderr: "timed out",
        timedOut: true,
        durationMs: request.timeoutMs,
      };
    }
    if (request.command === "pi-office-probe-output-cap") {
      return {
        exitCode: 0,
        stdout: "x".repeat(request.outputByteLimit + 1024),
        stderr: "",
        durationMs: 5,
      };
    }
    return {
      exitCode: 0,
      stdout: `ran:${request.command}`,
      stderr: "",
      touchedPaths: [request.cwd],
      durationMs: 5,
    };
  }
}

function testPaths() {
  const workspaceDir = resolve(process.cwd(), ".codex-office-shell-test", "workspace");
  const dataDir = resolve(process.cwd(), ".codex-office-shell-test", "data");
  return { workspaceDir, dataDir };
}

test("companion shell detection is disabled by default and degrades instead of enabling raw host shells", () => {
  const disabled = detectShellSandboxBackend({
    platform: "win32",
    env: {},
    commandExists: () => true,
  });
  assert.equal(disabled.state, "unavailable");
  assert.equal(disabled.backend, "none");
  assert.match(disabled.reason, /disabled by default/i);

  const dockerDetected = detectShellSandboxBackend({
    platform: "win32",
    env: { PI_OFFICE_ENABLE_SHELL_SANDBOX: "1" },
    commandExists: (command) => command === "docker",
  });
  assert.equal(dockerDetected.state, "degraded");
  assert.equal(dockerDetected.backend, "docker");
  assert.match(dockerDetected.reason, /destructive probes pass/i);

  const linuxMissing = detectShellSandboxBackend({
    platform: "linux",
    env: { PI_OFFICE_ENABLE_SHELL_SANDBOX: "1" },
    commandExists: () => false,
  });
  assert.equal(linuxMissing.state, "unavailable");
  assert.match(linuxMissing.reason, /raw host bash stays unavailable/i);
});

test("companion shell policy probes deny destructive, secret, network, package, and git operations", () => {
  const { workspaceDir, dataDir } = testPaths();
  const policy = createShellPolicy({ sessionId: "policy-probes", dataDir, workspaceDir });
  const probes = runShellPolicyProbes(policy);
  const ids = new Set(probes.map((probe) => probe.id));

  for (const expected of [
    "deny-write-outside-scratch",
    "deny-delete-outside-scratch",
    "deny-env-read",
    "deny-symlink-secret-escape",
    "deny-network",
    "deny-package-install",
    "deny-git-push",
    "deny-git-commit",
  ]) {
    assert.equal(ids.has(expected), true, `${expected} should be part of SECURITY-006 probes.`);
  }

  assert.ok(probes.every((probe) => probe.ok && probe.blocked), "all policy probes must be blocked before execution");
});

test("companion shell validation allows scratch writes only inside scratch and denies workspace mutation", () => {
  const { workspaceDir, dataDir } = testPaths();
  const policy = createShellPolicy({ sessionId: "scratch", dataDir, workspaceDir });
  assert.ok(policy.scratchRoot);

  const workspaceWrite = validateShellRequest(policy, {
    command: "echo bad > report.txt",
    cwd: workspaceDir,
    category: "read-only",
  });
  assert.equal(workspaceWrite.ok, false);
  assert.match(workspaceWrite.error ?? "", /Mutating shell commands require/i);

  const scratchWrite = validateShellRequest(policy, {
    command: "echo ok > result.txt",
    cwd: policy.scratchRoot,
    category: "scratch-write",
  });
  assert.equal(scratchWrite.ok, true);

  const scratchEscape = validateShellRequest(policy, {
    command: `echo bad > "${resolve(workspaceDir, "report.txt")}"`,
    cwd: policy.scratchRoot,
    category: "scratch-write",
  });
  assert.equal(scratchEscape.ok, false);
  assert.match(scratchEscape.error ?? "", /must stay inside/i);
});

test("companion shell execution stays unavailable unless detection and runner both pass", async () => {
  const { workspaceDir, dataDir } = testPaths();
  const runner = new FakeShellRunner();
  const sandbox = new CompanionShellSandbox({
    sessionId: "unavailable",
    dataDir,
    workspaceDir,
    detection: {
      state: "unavailable",
      backend: "none",
      reason: "disabled for test",
    },
    runner,
  });

  const capability = sandbox.getCapability();
  assert.equal(capability.state, "unavailable");

  const result = await sandbox.execute({ command: "ls", cwd: workspaceDir, category: "read-only" });
  assert.equal(result.ok, false);
  assert.equal(runner.calls.length, 0, "unavailable shell must not call the runner");
  assert.match(result.error ?? "", /disabled for test/);
});

test("companion shell runner receives scrubbed env, output caps, timeout probes, and BashOperations adapter", async () => {
  const { workspaceDir, dataDir } = testPaths();
  const runner = new FakeShellRunner();
  const sandbox = new CompanionShellSandbox({
    sessionId: "available",
    dataDir,
    workspaceDir,
    detection: {
      state: "available",
      backend: "test",
      reason: "test runner",
    },
    runner,
    env: {
      PATH: "/usr/bin",
      HOME: "/home/tester",
      OPENAI_API_KEY: "must-not-leak",
      NPM_TOKEN: "must-not-leak",
    },
  });

  const capability = sandbox.getCapability();
  assert.equal(capability.state, "available");
  assert.equal(capability.backend, "test");

  const result = await sandbox.execute({ command: "ls", cwd: workspaceDir, category: "read-only" });
  assert.equal(result.ok, true);
  assert.equal(result.stdout, "ran:ls");
  assert.equal(runner.calls[0]?.env.PATH, "/usr/bin");
  assert.equal(runner.calls[0]?.env.OPENAI_API_KEY, undefined);
  assert.equal(runner.calls[0]?.env.NPM_TOKEN, undefined);

  const executionProbes = await runShellExecutionProbes(sandbox);
  assert.ok(executionProbes.every((probe) => probe.ok), "timeout and output-cap probes must pass before shell can be trusted");

  const chunks: Buffer[] = [];
  const ops = createCompanionBashOperations(sandbox);
  const bashResult = await ops.exec("pwd", workspaceDir, {
    onData: (chunk) => chunks.push(chunk),
  });
  assert.equal(bashResult.exitCode, 0);
  assert.match(Buffer.concat(chunks).toString("utf8"), /ran:pwd/);
});
