import type { TSchema } from "@sinclair/typebox";
import type { OfficeHost, OfficeToolName, ToolCategory } from "@pi-office/pi-office-pack/protocol";

export type OfficeToolHostSupport = "all" | readonly OfficeHost[];
export type OfficeToolExecutorKind = "office-bridge" | "reviewable-word-edit" | "companion-native-capture";

export interface OfficeToolDefinition {
  name: OfficeToolName;
  hosts: OfficeToolHostSupport;
  category: ToolCategory;
  label: string;
  description: string;
  parameters: TSchema;
  executor: OfficeToolExecutorKind;
}

export function officeToolSupportsHost(definition: OfficeToolDefinition, host: OfficeHost): boolean {
  return definition.hosts === "all" || definition.hosts.includes(host);
}
