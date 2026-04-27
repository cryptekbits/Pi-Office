import type { TSchema } from "@sinclair/typebox";
import type { OfficeHost, OfficeToolName, ToolCategory } from "@pi-office/pi-office-pack/protocol";

export type OfficeToolHostSupport = "all" | readonly OfficeHost[];
export type OfficeToolExecutorKind = "office-bridge" | "reviewable-word-edit" | "companion-native-capture" | "runtime-registry";
export type OfficeToolDiscoveryVisibility = "core" | "deferred" | "conditional";
export type OfficeToolRuntimeOwner = "taskpane" | "companion" | "browser" | "office-js";
export type OfficeToolRiskLevel = "low" | "medium" | "high" | "critical" | "escape-hatch";
export interface OfficeToolRequirementSet {
  name: string;
  minVersion: string;
  note?: string | undefined;
}

export interface OfficeToolDefinition {
  name: OfficeToolName;
  hosts: OfficeToolHostSupport;
  category: ToolCategory;
  label: string;
  description: string;
  parameters: TSchema;
  executor: OfficeToolExecutorKind;
  core?: boolean | undefined;
  deferred?: boolean | undefined;
  exposure?: "core" | "deferred" | "conditional" | undefined;
  capabilityTags?: readonly string[] | undefined;
  compactSummary?: string | undefined;
  riskLevel?: OfficeToolRiskLevel | undefined;
  runtimeOwner?: OfficeToolRuntimeOwner | undefined;
  requirementSets?: readonly (OfficeToolRequirementSet | string)[] | undefined;
  capabilityIds?: readonly string[] | undefined;
  keywords?: readonly string[] | undefined;
  fallback?: string | undefined;
  requiresCompanion?: boolean | undefined;
  summary?: string | undefined;
  searchKeywords?: readonly string[] | undefined;
  discoveryId?: string | undefined;
  discoverability?: unknown;
  capabilityId?: string | undefined;
  risk?: OfficeToolRiskLevel | undefined;
  discovery?: {
    visibility?: OfficeToolDiscoveryVisibility | undefined;
    tier?: "core" | "specialized" | "escape-hatch" | string | undefined;
    capabilityId?: string | undefined;
    capabilityIds?: readonly string[] | undefined;
    keywords?: readonly string[] | undefined;
    requirementSets?: readonly (OfficeToolRequirementSet | string)[] | undefined;
    risk?: "low" | "medium" | "high" | "critical" | "escape-hatch" | undefined;
    riskLevel?: OfficeToolRiskLevel | undefined;
    summary?: string | undefined;
    fallback?: string | undefined;
  } | undefined;
}

export function officeToolSupportsHost(definition: OfficeToolDefinition, host: OfficeHost): boolean {
  return definition.hosts === "all" || definition.hosts.includes(host);
}
