import { readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { DEFAULT_COMPANION_HOST, DEFAULT_COMPANION_PORT } from "@pi-office/pi-office-pack";

const sourceDir = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(sourceDir, "../..");

export interface CompanionConfig {
  host: string;
  port: number;
  endpoint: string;
  identity: string;
  repoRoot: string;
  certDir: string;
  dataDir: string;
  tls: {
    pfx: Buffer;
    passphrase: string;
  };
}

export function loadConfig(): CompanionConfig {
  const host = process.env.PI_OFFICE_HOST ?? DEFAULT_COMPANION_HOST;
  const port = Number(process.env.PI_OFFICE_PORT ?? DEFAULT_COMPANION_PORT);
  const certDir = join(repoRoot, "addin", "certs");
  const pfxPath = join(certDir, "localhost.pfx");
  const passphrasePath = join(certDir, "passphrase.txt");
  const endpoint = `https://${host}:${port}`;

  return {
    host,
    port,
    endpoint,
    identity: `pi-office-companion@${host}:${port}`,
    repoRoot,
    certDir,
    dataDir: join(repoRoot, ".pi-office", "companion"),
    tls: {
      pfx: readFileSync(pfxPath),
      passphrase: readFileSync(passphrasePath, "utf8"),
    },
  };
}
