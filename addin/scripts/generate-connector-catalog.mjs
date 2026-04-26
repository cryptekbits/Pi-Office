import { readFileSync, writeFileSync } from "node:fs";
import { dirname, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { parse } from "yaml";

const __dirname = dirname(fileURLToPath(import.meta.url));
const addinRoot = resolve(__dirname, "..");
const yamlPath = resolve(addinRoot, "packages/pi-office-pack/src/connector-catalog.yaml");
const outputPath = resolve(addinRoot, "apps/taskpane/src/lib/runtime/connector-catalog-generated.ts");
const checkOnly = process.argv.includes("--check");

const VALID_OFFICIALNESS = new Set([
  "official",
  "official_preview",
  "community",
  "provider_reference",
  "deprecated",
  "experimental",
  "planned",
]);
const VALID_AVAILABILITY = new Set(["available", "needs_companion", "planned", "advanced"]);
const VALID_BROWSER_DIRECT = new Set(["supported", "unsupported", "unknown"]);

function asObject(value, label) {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    throw new Error(`${label} must be an object.`);
  }
  return value;
}

function asArray(value, label) {
  if (!Array.isArray(value)) {
    throw new Error(`${label} must be an array.`);
  }
  return value;
}

function asString(value, label) {
  if (typeof value !== "string" || !value.trim()) {
    throw new Error(`${label} must be a non-empty string.`);
  }
  return value.trim();
}

function asBoolean(value, label) {
  if (typeof value !== "boolean") {
    throw new Error(`${label} must be a boolean.`);
  }
  return value;
}

function validateUrl(value, label) {
  const url = asString(value, label);
  try {
    const parsed = new URL(url);
    if (parsed.protocol !== "https:" && parsed.protocol !== "http:") {
      throw new Error("unsupported protocol");
    }
  } catch {
    throw new Error(`${label} must be an HTTP(S) URL.`);
  }
  return url;
}

function compact(value) {
  if (Array.isArray(value)) {
    return value.map(compact).filter((entry) => entry !== undefined);
  }
  if (value && typeof value === "object") {
    const next = {};
    for (const [key, entry] of Object.entries(value)) {
      const normalized = compact(entry);
      if (normalized !== undefined) {
        next[key] = normalized;
      }
    }
    return next;
  }
  return value === undefined ? undefined : value;
}

function validateProfile(connector, profile, index, errors) {
  const label = `${connector.id}.setupProfiles[${index}]`;
  asString(profile.id, `${label}.id`);
  asString(profile.label, `${label}.label`);
  asString(profile.transport, `${label}.transport`);
  asString(profile.authMethod, `${label}.authMethod`);
  asBoolean(profile.requiresCompanion, `${label}.requiresCompanion`);

  const officialness = asString(profile.officialness, `${label}.officialness`);
  if (!VALID_OFFICIALNESS.has(officialness)) {
    errors.push(`${label}.officialness must be one of ${[...VALID_OFFICIALNESS].join(", ")}.`);
  }
  const availability = asString(profile.availability, `${label}.availability`);
  if (!VALID_AVAILABILITY.has(availability)) {
    errors.push(`${label}.availability must be one of ${[...VALID_AVAILABILITY].join(", ")}.`);
  }
  const browserDirect = asString(profile.browserDirect, `${label}.browserDirect`);
  if (!VALID_BROWSER_DIRECT.has(browserDirect)) {
    errors.push(`${label}.browserDirect must be one of ${[...VALID_BROWSER_DIRECT].join(", ")}.`);
  }

  validateUrl(profile.docsUrl, `${label}.docsUrl`);
  validateUrl(profile.endpointEvidenceUrl, `${label}.endpointEvidenceUrl`);
  if (profile.authMethod !== "none") {
    validateUrl(profile.authEvidenceUrl, `${label}.authEvidenceUrl`);
  }
  asString(profile.checkedAt, `${label}.checkedAt`);
  asArray(profile.riskNotes, `${label}.riskNotes`);

  if (profile.officialness === "planned" && profile.setupDisabled !== true) {
    errors.push(`${label} is planned but setupDisabled is not true.`);
  }
  if (profile.transport === "local_stdio" && profile.requiresCompanion !== true) {
    errors.push(`${label} is local stdio but does not require the companion.`);
  }
  if (profile.browserDirect === "supported" && profile.requiresCompanion === true) {
    errors.push(`${label} is browserDirect=supported but still requires companion.`);
  }
}

function normalizeCatalog(raw) {
  const root = asObject(raw, "connector catalog YAML");
  const connectors = asArray(root.connectors, "connectors");
  const errors = [];
  const seen = new Set();
  const normalized = connectors.map((connectorRaw, connectorIndex) => {
    const connector = asObject(connectorRaw, `connectors[${connectorIndex}]`);
    const id = asString(connector.id, `connectors[${connectorIndex}].id`);
    if (seen.has(id)) errors.push(`Duplicate connector id ${id}.`);
    seen.add(id);
    asString(connector.name, `${id}.name`);
    asString(connector.vendor, `${id}.vendor`);
    validateUrl(connector.docsUrl, `${id}.docsUrl`);
    const profiles = asArray(connector.setupProfiles, `${id}.setupProfiles`);
    profiles.forEach((profileRaw, index) => validateProfile(connector, asObject(profileRaw, `${id}.setupProfiles[${index}]`), index, errors));
    return compact(connector);
  });

  if (errors.length) {
    throw new Error(`Connector catalog validation failed:\n- ${errors.join("\n- ")}`);
  }
  return normalized;
}

function literal(value) {
  return JSON.stringify(value, null, 2);
}

const parsed = parse(readFileSync(yamlPath, "utf8"));
const catalog = normalizeCatalog(parsed);
const output = `// Generated by ${relative(addinRoot, fileURLToPath(import.meta.url)).replaceAll("\\", "/")}.
// Do not edit directly. Edit packages/pi-office-pack/src/connector-catalog.yaml instead.

import type { ConnectorCatalogItem } from "@pi-office/pi-office-pack";

export const CONNECTOR_CATALOG_REVISION = ${JSON.stringify(parsed.revision ?? "unknown")};

export const CONNECTOR_CATALOG = ${literal(catalog)} as ConnectorCatalogItem[];
`;

if (checkOnly) {
  const existing = readFileSync(outputPath, "utf8");
  if (existing !== output) {
    throw new Error("connector-catalog-generated.ts is out of date. Run npm run generate:connector-catalog.");
  }
} else {
  writeFileSync(outputPath, output);
}
