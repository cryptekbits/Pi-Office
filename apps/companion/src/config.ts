import { readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { DEFAULT_COMPANION_HOST, DEFAULT_COMPANION_PORT } from "@pi-office/pi-office-pack";

const sourceDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(sourceDir, "../../..");

export interface CompanionConfig {
  host: string;
  port: number;
  origin: string;
  mode: "development" | "production";
  repoRoot: string;
  dataDir: string;
  agentDir: string;
  scratchDir: string;
  connectorsStatePath: string;
  connectorsMcpConfigPath: string;
  connectorsRuntimeDir: string;
  connectorsSecretsDir: string;
  connectorsVaultPath: string;
  taskpaneRoot: string;
  taskpaneDist: string;
  tls: {
    pfx: Buffer;
    passphrase: string;
  };
}

export function loadConfig(): CompanionConfig {
  const host = process.env.PI_OFFICE_HOST ?? DEFAULT_COMPANION_HOST;
  const port = Number(process.env.PI_OFFICE_PORT ?? DEFAULT_COMPANION_PORT);
  const mode = process.env.NODE_ENV === "production" ? "production" : "development";
  const certDir = join(repoRoot, "certs");
  const pfxPath = join(certDir, "localhost.pfx");
  const passphrasePath = join(certDir, "passphrase.txt");

  return {
    host,
    port,
    origin: `https://${host}:${port}`,
    mode,
    repoRoot,
    dataDir: join(repoRoot, ".pi-office"),
    agentDir: join(repoRoot, ".pi-office", "agent"),
    scratchDir: join(repoRoot, ".pi-office", "scratch"),
    connectorsStatePath: join(repoRoot, ".pi-office", "connectors.json"),
    connectorsMcpConfigPath: join(repoRoot, ".pi-office", "agent", "mcp.json"),
    connectorsRuntimeDir: join(repoRoot, ".pi-office", "agent", "connectors"),
    connectorsSecretsDir: join(repoRoot, ".pi-office", "secrets"),
    connectorsVaultPath: join(repoRoot, ".pi-office", "secrets", "connector-vault.json"),
    taskpaneRoot: join(repoRoot, "apps", "taskpane"),
    taskpaneDist: join(repoRoot, "apps", "taskpane", "dist"),
    tls: {
      pfx: readFileSync(pfxPath),
      passphrase: readFileSync(passphrasePath, "utf8"),
    },
  };
}
