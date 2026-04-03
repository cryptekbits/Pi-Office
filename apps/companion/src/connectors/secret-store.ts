import { createCipheriv, createDecipheriv, randomBytes, scryptSync } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, unlinkSync, writeFileSync } from "node:fs";
import { dirname } from "node:path";
import { spawnSync } from "node:child_process";
import { hostname, userInfo } from "node:os";

export interface StoredSecretReference {
  backend: "windows-dpapi" | "encrypted-file";
  key: string;
}

interface EncryptedVaultEntry {
  iv: string;
  tag: string;
  data: string;
}

interface EncryptedVaultFile {
  version: 1;
  salt: string;
  entries: Record<string, EncryptedVaultEntry>;
}

function ensureDir(filePath: string): void {
  mkdirSync(dirname(filePath), { recursive: true });
}

function readJsonFile<T>(filePath: string, fallback: T): T {
  if (!existsSync(filePath)) {
    return fallback;
  }

  try {
    return JSON.parse(readFileSync(filePath, "utf8")) as T;
  } catch {
    return fallback;
  }
}

function writeJsonFile(filePath: string, value: unknown): void {
  ensureDir(filePath);
  writeFileSync(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

function machineFingerprint(): string {
  return `${process.platform}|${process.arch}|${hostname()}|${userInfo().username}`;
}

function powershellProtect(plaintext: string): string | undefined {
  const payload = Buffer.from(plaintext, "utf8").toString("base64");
  const script = [
    "$bytes = [Convert]::FromBase64String($args[0])",
    "$protected = [Security.Cryptography.ProtectedData]::Protect($bytes, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)",
    "[Convert]::ToBase64String($protected)",
  ].join("; ");
  const result = spawnSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script, payload], {
    encoding: "utf8",
    windowsHide: true,
  });
  if (result.status !== 0) {
    return undefined;
  }
  const output = result.stdout.trim();
  return output || undefined;
}

function powershellUnprotect(ciphertext: string): string | undefined {
  const script = [
    "$bytes = [Convert]::FromBase64String($args[0])",
    "$plain = [Security.Cryptography.ProtectedData]::Unprotect($bytes, $null, [Security.Cryptography.DataProtectionScope]::CurrentUser)",
    "[Text.Encoding]::UTF8.GetString($plain)",
  ].join("; ");
  const result = spawnSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-Command", script, ciphertext], {
    encoding: "utf8",
    windowsHide: true,
  });
  if (result.status !== 0) {
    return undefined;
  }
  const output = result.stdout;
  return typeof output === "string" ? output.replace(/\r?\n$/, "") : undefined;
}

function createFallbackKey(salt: Buffer): Buffer {
  return scryptSync(machineFingerprint(), salt, 32);
}

function encryptFallbackSecret(secret: string, salt: Buffer): EncryptedVaultEntry {
  const iv = randomBytes(12);
  const key = createFallbackKey(salt);
  const cipher = createCipheriv("aes-256-gcm", key, iv);
  const encrypted = Buffer.concat([cipher.update(secret, "utf8"), cipher.final()]);
  const tag = cipher.getAuthTag();

  return {
    iv: iv.toString("base64"),
    tag: tag.toString("base64"),
    data: encrypted.toString("base64"),
  };
}

function decryptFallbackSecret(entry: EncryptedVaultEntry, salt: Buffer): string | undefined {
  try {
    const key = createFallbackKey(salt);
    const decipher = createDecipheriv("aes-256-gcm", key, Buffer.from(entry.iv, "base64"));
    decipher.setAuthTag(Buffer.from(entry.tag, "base64"));
    const decrypted = Buffer.concat([
      decipher.update(Buffer.from(entry.data, "base64")),
      decipher.final(),
    ]);
    return decrypted.toString("utf8");
  } catch {
    return undefined;
  }
}

export class ConnectorSecretStore {
  private readonly dpapiAvailable: boolean;

  constructor(
    private readonly secretsDir: string,
    private readonly vaultPath: string,
  ) {
    this.dpapiAvailable = this.checkDpapiAvailability();
  }

  isAvailable(): boolean {
    return true;
  }

  store(secret: string, keyHint: string): StoredSecretReference {
    const sanitizedKey = keyHint.replace(/[^a-zA-Z0-9_-]+/g, "-").replace(/^-+|-+$/g, "") || "connector-secret";
    if (this.dpapiAvailable) {
      const encrypted = powershellProtect(secret);
      if (encrypted) {
        const filePath = this.dpapiPathForKey(sanitizedKey);
        ensureDir(filePath);
        writeFileSync(filePath, encrypted, "utf8");
        return { backend: "windows-dpapi", key: sanitizedKey };
      }
    }

    const vault = this.readFallbackVault();
    const salt = Buffer.from(vault.salt, "base64");
    vault.entries[sanitizedKey] = encryptFallbackSecret(secret, salt);
    this.writeFallbackVault(vault);
    return { backend: "encrypted-file", key: sanitizedKey };
  }

  read(reference: StoredSecretReference | undefined): string | undefined {
    if (!reference) {
      return undefined;
    }

    if (reference.backend === "windows-dpapi") {
      const filePath = this.dpapiPathForKey(reference.key);
      if (!existsSync(filePath)) {
        return undefined;
      }
      return powershellUnprotect(readFileSync(filePath, "utf8").trim());
    }

    const vault = this.readFallbackVault();
    const entry = vault.entries[reference.key];
    if (!entry) {
      return undefined;
    }

    return decryptFallbackSecret(entry, Buffer.from(vault.salt, "base64"));
  }

  remove(reference: StoredSecretReference | undefined): void {
    if (!reference) {
      return;
    }

    if (reference.backend === "windows-dpapi") {
      const filePath = this.dpapiPathForKey(reference.key);
      if (existsSync(filePath)) {
        unlinkSync(filePath);
      }
      return;
    }

    const vault = this.readFallbackVault();
    if (vault.entries[reference.key]) {
      delete vault.entries[reference.key];
      this.writeFallbackVault(vault);
    }
  }

  private checkDpapiAvailability(): boolean {
    if (process.platform !== "win32") {
      return false;
    }

    const protectedValue = powershellProtect("pi-office-secret-store-probe");
    if (!protectedValue) {
      return false;
    }

    return powershellUnprotect(protectedValue) === "pi-office-secret-store-probe";
  }

  private dpapiPathForKey(key: string): string {
    return `${this.secretsDir}\\${key}.secret`;
  }

  private readFallbackVault(): EncryptedVaultFile {
    const existing = readJsonFile<EncryptedVaultFile | null>(this.vaultPath, null);
    if (existing && existing.version === 1 && typeof existing.salt === "string" && existing.entries) {
      return existing;
    }

    return {
      version: 1,
      salt: randomBytes(16).toString("base64"),
      entries: {},
    };
  }

  private writeFallbackVault(vault: EncryptedVaultFile): void {
    writeJsonFile(this.vaultPath, vault);
  }
}
