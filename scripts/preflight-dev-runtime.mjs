import { spawnSync } from "node:child_process";
import { existsSync, readFileSync } from "node:fs";
import net from "node:net";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = dirname(fileURLToPath(new URL("../package.json", import.meta.url)));
const certDir = join(repoRoot, "certs");
const pfxPath = join(certDir, "localhost.pfx");
const passphrasePath = join(certDir, "passphrase.txt");
const thumbprintPath = join(certDir, "thumbprint.txt");
const requiredPort = 3443;

function fail(message) {
  console.error(`[preflight] ${message}`);
  process.exit(1);
}

function runPowerShell(command) {
  const candidates = [
    "C:\\Program Files\\PowerShell\\7\\pwsh.exe",
    "pwsh",
    "powershell.exe",
  ];

  for (const candidate of candidates) {
    const result = spawnSync(candidate, ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command], {
      encoding: "utf8",
      stdio: "pipe",
    });

    if (result.error?.code === "ENOENT") {
      continue;
    }
    return result;
  }

  return null;
}

function ensureCertFiles() {
  if (!existsSync(pfxPath) || !existsSync(passphrasePath)) {
    fail(
      `Missing dev certificate files in ${certDir}. Run "npm run prepare:certs" and retry.`,
    );
  }

  const passphrase = readFileSync(passphrasePath, "utf8").trim();
  if (!passphrase) {
    fail(`Certificate passphrase file is empty: ${passphrasePath}`);
  }
}

function ensureTrustedCertOnWindows() {
  if (process.platform !== "win32") return;
  if (!existsSync(thumbprintPath)) {
    fail(`Missing certificate thumbprint file: ${thumbprintPath}. Run "npm run prepare:certs".`);
  }

  const thumbprint = readFileSync(thumbprintPath, "utf8").trim().toUpperCase();
  if (!thumbprint) {
    fail(`Certificate thumbprint is empty: ${thumbprintPath}`);
  }

  const command = `
$thumb="${thumbprint}"
$found = Get-ChildItem Cert:\\CurrentUser\\Root | Where-Object { $_.Thumbprint -eq $thumb } | Select-Object -First 1
if ($found) { Write-Output "ok"; exit 0 } else { Write-Output "missing"; exit 2 }
`;

  const result = runPowerShell(command);
  if (!result) {
    fail("PowerShell was not found; cannot verify trusted localhost certificate.");
  }
  if (result.status !== 0) {
    fail(
      `Trusted localhost certificate not found in Cert:\\CurrentUser\\Root. Run "npm run prepare:certs" and trust the generated cert.`,
    );
  }
}

function ensurePortAvailable(port) {
  return new Promise((resolve, reject) => {
    const server = net.createServer();
    server.unref();

    server.once("error", (error) => {
      if (error && typeof error === "object" && "code" in error && error.code === "EADDRINUSE") {
        reject(new Error(`Port ${port} is already in use. Stop the existing process or free the port before starting dev.`));
        return;
      }
      reject(error instanceof Error ? error : new Error(String(error)));
    });

    server.listen(port, "127.0.0.1", () => {
      server.close((closeError) => {
        if (closeError) {
          reject(closeError);
          return;
        }
        resolve();
      });
    });
  });
}

async function main() {
  ensureCertFiles();
  ensureTrustedCertOnWindows();
  await ensurePortAvailable(requiredPort);
  console.log(`[preflight] OK (certs present/trusted, port ${requiredPort} available).`);
}

await main().catch((error) => {
  fail(error instanceof Error ? error.message : String(error));
});
