import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = dirname(fileURLToPath(new URL("../package.json", import.meta.url)));
const certDir = join(repoRoot, "certs");
const pfxPath = join(certDir, "localhost.pfx");
const passphrasePath = join(certDir, "passphrase.txt");
const thumbprintPath = join(certDir, "thumbprint.txt");

if (existsSync(pfxPath) && existsSync(passphrasePath)) {
  console.log(`Using existing dev certificate at ${pfxPath}`);
  process.exit(0);
}

mkdirSync(certDir, { recursive: true });

const passphrase = "pi-office-localhost";
const script = `
$ErrorActionPreference = "Stop"
$certPath = "${pfxPath.replace(/\\/g, "\\\\")}"
$passPath = "${passphrasePath.replace(/\\/g, "\\\\")}"
$thumbPath = "${thumbprintPath.replace(/\\/g, "\\\\")}"
$cerPath = [System.IO.Path]::ChangeExtension($certPath, ".cer")
$password = New-Object System.Security.SecureString
"${passphrase}".ToCharArray() | ForEach-Object { $password.AppendChar($_) }
$password.MakeReadOnly()
$existing = Get-ChildItem Cert:\\CurrentUser\\My | Where-Object { $_.Subject -eq "CN=localhost" } | Select-Object -First 1
if (-not $existing) {
  $existing = New-SelfSignedCertificate -DnsName "localhost" -CertStoreLocation "Cert:\\CurrentUser\\My" -FriendlyName "Pi-Office localhost" -NotAfter (Get-Date).AddYears(5) -KeyExportPolicy Exportable -KeyAlgorithm RSA -KeyLength 2048 -HashAlgorithm SHA256
}
Export-Certificate -Cert $existing -FilePath $cerPath -Force | Out-Null
Import-Certificate -FilePath $cerPath -CertStoreLocation "Cert:\\CurrentUser\\Root" | Out-Null
Export-PfxCertificate -Cert $existing -FilePath $certPath -Password $password | Out-Null
Set-Content -LiteralPath $passPath -Value "${passphrase}" -NoNewline
Set-Content -LiteralPath $thumbPath -Value $existing.Thumbprint -NoNewline
Write-Output "Created localhost certificate: $certPath"
`;

const powerShellCandidates = [
  "C:\\Program Files\\PowerShell\\7\\pwsh.exe",
  "pwsh",
  "powershell.exe",
];

let result = null;
for (const candidate of powerShellCandidates) {
  const attempt = spawnSync(candidate, ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script], {
    stdio: "inherit",
  });
  if (attempt.error?.code === "ENOENT") {
    continue;
  }
  result = attempt;
  break;
}

if (!result) {
  console.error("PowerShell was not found. Install PowerShell 7 or ensure powershell.exe is available in PATH.");
  process.exit(1);
}

if (result.status !== 0) {
  process.exit(result.status ?? 1);
}
