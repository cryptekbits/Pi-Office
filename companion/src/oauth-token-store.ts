import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import type { CompanionConfig } from "./config.js";

export interface OAuthTokenPersistence {
  readonly storageKind: "windows-dpapi" | "local-json-unsupported";
  readonly secure: boolean;
  readonly storagePath: string;
  readonly legacyPath?: string | undefined;
  load(): string | undefined;
  save(value: string): void;
  clear(): void;
}

interface WindowsDpapiOptions {
  protect?: ((value: string) => string) | undefined;
  unprotect?: ((value: string) => string) | undefined;
}

interface CreateStoreOptions extends WindowsDpapiOptions {
  platform?: NodeJS.Platform | undefined;
}

function runDpapiPowerShell(mode: "protect" | "unprotect", value: string): string {
  const script = `
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
$inputText = [Console]::In.ReadToEnd()
$decoded = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($inputText))
if ("${mode}" -eq "protect") {
  $secure = ConvertTo-SecureString -String $decoded -AsPlainText -Force
  [Console]::Out.Write((ConvertFrom-SecureString -SecureString $secure))
} else {
  $secure = ConvertTo-SecureString -String $decoded
  $ptr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
  try {
    $plain = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    [Console]::Out.Write([Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($plain)))
  } finally {
    if ($ptr -ne [IntPtr]::Zero) {
      [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
    }
  }
}
`;
  // Feed the secret to PowerShell after the script has been compiled, keeping it out of argv.
  const encodedInput = Buffer.from(value, "utf8").toString("base64");
  const output = execFileSync("powershell.exe", [
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-EncodedCommand",
    Buffer.from(script, "utf16le").toString("base64"),
  ], {
    input: encodedInput,
    encoding: "utf8",
    env: {
      SystemRoot: process.env.SystemRoot,
      PATH: process.env.PATH,
      PATHEXT: process.env.PATHEXT,
    },
  });
  const trimmed = output.trim();
  return mode === "protect" ? trimmed : Buffer.from(trimmed, "base64").toString("utf8");
}

function protectWithWindowsDpapi(value: string): string {
  return runDpapiPowerShell("protect", value);
}

function unprotectWithWindowsDpapi(value: string): string {
  return runDpapiPowerShell("unprotect", value);
}

function ensureParent(path: string): void {
  mkdirSync(dirname(path), { recursive: true });
}

class LocalJsonUnsupportedTokenStore implements OAuthTokenPersistence {
  readonly storageKind = "local-json-unsupported" as const;
  readonly secure = false;
  readonly legacyPath = undefined;

  constructor(readonly storagePath: string) {}

  load(): string | undefined {
    return existsSync(this.storagePath) ? readFileSync(this.storagePath, "utf8") : undefined;
  }

  save(value: string): void {
    ensureParent(this.storagePath);
    writeFileSync(this.storagePath, value, { mode: 0o600 });
  }

  clear(): void {
    rmSync(this.storagePath, { force: true });
  }
}

class WindowsDpapiTokenStore implements OAuthTokenPersistence {
  readonly storageKind = "windows-dpapi" as const;
  readonly secure = true;

  constructor(
    readonly storagePath: string,
    readonly legacyPath: string,
    private readonly protect: (value: string) => string,
    private readonly unprotect: (value: string) => string,
  ) {}

  load(): string | undefined {
    if (existsSync(this.storagePath)) {
      const envelope = JSON.parse(readFileSync(this.storagePath, "utf8")) as { protection?: string; payload?: string };
      if (envelope.protection !== "windows-dpapi-current-user" || !envelope.payload) {
        throw new Error("Companion OAuth token store has an unsupported encrypted envelope.");
      }
      return this.unprotect(envelope.payload);
    }

    if (!existsSync(this.legacyPath)) {
      return undefined;
    }

    const legacy = readFileSync(this.legacyPath, "utf8");
    this.save(legacy);
    renameSync(this.legacyPath, `${this.legacyPath}.migrated`);
    return legacy;
  }

  save(value: string): void {
    ensureParent(this.storagePath);
    const envelope = {
      version: 1,
      protection: "windows-dpapi-current-user",
      payload: this.protect(value),
    };
    writeFileSync(this.storagePath, `${JSON.stringify(envelope, null, 2)}\n`, { mode: 0o600 });
  }

  clear(): void {
    rmSync(this.storagePath, { force: true });
    rmSync(this.legacyPath, { force: true });
    rmSync(`${this.legacyPath}.migrated`, { force: true });
  }
}

export function createCompanionOAuthTokenStore(
  config: CompanionConfig,
  options: CreateStoreOptions = {},
): OAuthTokenPersistence {
  const legacyPath = join(config.dataDir, "connector-oauth-tokens.json");
  if ((options.platform ?? process.platform) === "win32") {
    return new WindowsDpapiTokenStore(
      join(config.dataDir, "connector-oauth-tokens.dpapi.json"),
      legacyPath,
      options.protect ?? protectWithWindowsDpapi,
      options.unprotect ?? unprotectWithWindowsDpapi,
    );
  }
  return new LocalJsonUnsupportedTokenStore(legacyPath);
}
