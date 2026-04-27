import type {
  OfficeHost,
  OfficeToolName,
  ToolCapabilityDetail,
  ToolCapabilityRequirement,
  ToolCapabilityRisk,
  ToolCapabilityRuntime,
  ToolCapabilitySearchResult,
  ToolCapabilitySupportStatus,
  ToolCategory,
} from "@pi-office/pi-office-pack/protocol";
import { COMMON_OFFICE_TOOL_DEFINITIONS } from "./common";
import { EXCEL_OFFICE_TOOL_DEFINITIONS } from "./excel";
import { POWERPOINT_OFFICE_TOOL_DEFINITIONS } from "./powerpoint";
import type { OfficeToolDefinition } from "./types";
import { officeToolSupportsHost } from "./types";
import { WORD_OFFICE_TOOL_DEFINITIONS } from "./word";

export type {
  OfficeToolDefinition,
  OfficeToolExecutorKind,
  OfficeToolHostSupport,
  OfficeToolRequirementSet,
  OfficeToolRiskLevel,
  OfficeToolRuntimeOwner,
} from "./types";
export { officeToolSupportsHost } from "./types";

export const OFFICE_TOOL_DEFINITIONS: readonly OfficeToolDefinition[] = [
  ...COMMON_OFFICE_TOOL_DEFINITIONS,
  ...WORD_OFFICE_TOOL_DEFINITIONS,
  ...EXCEL_OFFICE_TOOL_DEFINITIONS,
  ...POWERPOINT_OFFICE_TOOL_DEFINITIONS,
];

const officeToolDefinitionsByName = new Map<OfficeToolName, OfficeToolDefinition>(
  OFFICE_TOOL_DEFINITIONS.map((definition) => [definition.name, definition]),
);

export function getOfficeToolDefinition(name: OfficeToolName): OfficeToolDefinition {
  const definition = officeToolDefinitionsByName.get(name);
  if (!definition) {
    throw new Error(`Missing Office tool definition for ${name}.`);
  }
  return definition;
}

export function getOfficeToolDefinitionsForHost(host: OfficeHost): OfficeToolDefinition[] {
  return OFFICE_TOOL_DEFINITIONS.filter((definition) => officeToolSupportsHost(definition, host));
}

export function getCoreOfficeToolDefinitionsForHost(host: OfficeHost): OfficeToolDefinition[] {
  return getOfficeToolDefinitionsForHost(host).filter((definition) => isCoreOfficeToolDefinition(definition));
}

export function getDeferredOfficeToolDefinitionsForHost(host: OfficeHost): OfficeToolDefinition[] {
  return getOfficeToolDefinitionsForHost(host).filter((definition) => !isCoreOfficeToolDefinition(definition));
}

function isCoreOfficeToolDefinition(definition: OfficeToolDefinition): boolean {
  if (definition.core === true) return true;
  if (definition.deferred === true) return false;
  const visibility = definition.discovery?.visibility ?? definition.discovery?.tier;
  if (visibility === "deferred" || visibility === "specialized") return false;
  if (definition.name === "office_propose_edits" || definition.name === "office_execute_js") return false;
  return true;
}

function schemaPropertyNames(definition: OfficeToolDefinition): string[] {
  const properties = (definition.parameters as { properties?: unknown }).properties;
  return properties && typeof properties === "object" && !Array.isArray(properties)
    ? Object.keys(properties)
    : [];
}

export function toToolCapabilitySearchResult(
  definition: OfficeToolDefinition,
  host: OfficeHost,
  includeSchema = false,
): ToolCapabilitySearchResult | ToolCapabilityDetail {
  const risk = normalizeRisk(definition.riskLevel);
  const runtime = normalizeRuntime(definition.runtimeOwner, definition.executor);
  const requirements = normalizeRequirements(definition.requirementSets);
  const result: ToolCapabilitySearchResult = {
    id: definition.capabilityIds?.[0] ?? definition.name,
    toolName: definition.name,
    label: definition.label,
    description: definition.description,
    host,
    category: definition.category,
    runtime,
    risk,
    support: definition.requiresCompanion ? "companion_required" : requirementSupport(requirements),
    requirements,
    keywords: [...new Set([
      ...(definition.keywords ?? []),
      ...(definition.capabilityTags ?? []),
      definition.name,
      definition.label,
    ].map((entry) => entry.toLowerCase()))],
    source: "office-tool-registry",
    schemaAvailable: true,
    fallback: definition.fallback,
  };
  return includeSchema ? { ...result, parameters: definition.parameters } : result;
}

function normalizeRisk(value: OfficeToolDefinition["riskLevel"]): ToolCapabilityRisk {
  if (value === "critical" || value === "escape-hatch") return "escape_hatch";
  if (value === "high" || value === "medium" || value === "low") return value;
  return "medium";
}

function normalizeRuntime(
  runtimeOwner: OfficeToolDefinition["runtimeOwner"],
  executor: OfficeToolDefinition["executor"],
): ToolCapabilityRuntime {
  if (runtimeOwner === "companion" || executor === "companion-native-capture") return "companion";
  if (runtimeOwner === "browser") return "browser";
  if (runtimeOwner === "office-js") return "office-js";
  return "taskpane";
}

function normalizeRequirements(
  entries: OfficeToolDefinition["requirementSets"],
): ToolCapabilityRequirement[] | undefined {
  const requirements = (entries ?? []).map((entry): ToolCapabilityRequirement => {
    if (typeof entry === "string") {
      const match = entry.match(/^([A-Za-z]+(?:Desktop|Online|HiddenDocument)?)(?:\s+)?(.+)?$/);
      return {
        name: match?.[1] ?? entry,
        version: match?.[2],
      };
    }
    const maybeVersion = "version" in entry ? entry.version : undefined;
    return {
      name: entry.name,
      version: entry.minVersion ?? maybeVersion,
      note: entry.note,
      status: /desktop/i.test(entry.name) ? "desktop_only" : undefined,
    };
  });
  return requirements.length ? requirements : undefined;
}

function requirementSupport(requirements: ToolCapabilityRequirement[] | undefined): ToolCapabilitySupportStatus {
  if (!requirements?.length) return "supported";
  if (requirements.some((entry) => entry.status === "desktop_only" || /desktop/i.test(entry.name))) return "desktop_only";
  if (requirements.some((entry) => /preview/i.test(entry.note ?? ""))) return "preview";
  return "supported";
}

function scoreTool(definition: OfficeToolDefinition, queryTerms: string[]): number {
  if (!queryTerms.length) return definition.deferred ? 5 : 8;
  const exactPhrase = queryTerms.join(" ");
  const haystack = [
    definition.name,
    definition.label,
    definition.description,
    definition.compactSummary,
    ...(definition.keywords ?? []),
    ...(definition.capabilityTags ?? []),
    ...(definition.capabilityIds ?? []),
  ].join(" ").toLowerCase();
  let score = haystack.includes(exactPhrase) ? 10 : 0;
  score += queryTerms.reduce((sum, term) => sum + (haystack.includes(term) ? 3 : 0), 0);
  if (definition.name.includes(queryTerms[0] ?? "\0")) score += 4;
  if (definition.deferred === true || !isCoreOfficeToolDefinition(definition)) score += 1;
  return score;
}

function queryTerms(query: string | undefined): string[] {
  return String(query ?? "")
    .toLowerCase()
    .split(/[^a-z0-9_:-]+/g)
    .map((term) => term.trim())
    .filter((term) => term.length > 1);
}

export interface OfficeToolDefinitionSearchInput {
  host: OfficeHost;
  query?: string | undefined;
  category?: ToolCategory | undefined;
  limit?: number | undefined;
  includeSchemas?: boolean | undefined;
}

export function searchOfficeToolDefinitions(input: OfficeToolDefinitionSearchInput): ToolCapabilitySearchResult[] {
  const terms = queryTerms(input.query);
  const limit = Math.max(1, Math.min(20, Math.trunc(input.limit ?? 5)));
  return getOfficeToolDefinitionsForHost(input.host)
    .filter((definition) => !input.category || definition.category === input.category)
    .map((definition) => ({
      definition,
      score: scoreTool(definition, terms),
    }))
    .filter((entry) => !terms.length || entry.score > 0)
    .sort((left, right) => right.score - left.score || left.definition.name.localeCompare(right.definition.name))
    .slice(0, limit)
    .map((entry) => ({
      ...toToolCapabilitySearchResult(entry.definition, input.host, input.includeSchemas),
      score: entry.score,
    } as ToolCapabilitySearchResult));
}

export function getOfficeToolCapabilityDetail(
  nameOrId: string,
  host: OfficeHost,
): ToolCapabilityDetail {
  const definition = OFFICE_TOOL_DEFINITIONS.find((entry) =>
    entry.name === nameOrId || entry.capabilityIds?.includes(nameOrId),
  );
  if (!definition) {
    throw new Error(`Unknown Office tool or capability: ${nameOrId}.`);
  }
  if (!officeToolSupportsHost(definition, host)) {
    throw new Error(`${definition.name} is not available for ${host}.`);
  }
  return toToolCapabilitySearchResult(definition, host, true) as ToolCapabilityDetail;
}
