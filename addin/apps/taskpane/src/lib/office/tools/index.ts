import type { OfficeHost, OfficeToolName } from "@pi-office/pi-office-pack/protocol";
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
