import type { CompanionState, OfficeHost } from "./protocol.js";

export type CapabilityOwner = "addin-only" | "companion-only" | "either";
export type CapabilityRuntime = "addin" | "companion";
export type CapabilityStatus = "available" | "unavailable";
export type CapabilityGroup = "works_without_companion" | "enhanced_by_companion" | "requires_companion";

export type CapabilityId =
  | "office_context"
  | "office_write"
  | "office_review"
  | "office_visual_snapshot"
  | "native_viewport_capture"
  | "powerpoint_native_visual"
  | "local_files"
  | "mcp_connectors"
  | "shell_sandbox"
  | "inference"
  | "provider_auth"
  | "image_generation"
  | "diagnostics"
  | "memory"
  | "background_jobs";

export interface CapabilityDefinition {
  id: CapabilityId;
  label: string;
  owner: CapabilityOwner;
  tools: string[];
  requiresCompanion: boolean;
  requiresSavedDocument: boolean;
  honestyLabel: string;
  group: CapabilityGroup;
}

export interface CapabilityResolution extends CapabilityDefinition {
  available: boolean;
  preferredRuntime: CapabilityRuntime;
  activeRuntime?: CapabilityRuntime | undefined;
  fallbackRuntime?: CapabilityRuntime | undefined;
  reason?: string | undefined;
}

export interface CapabilityResolutionInput {
  host: OfficeHost;
  documentSaved: boolean;
  companion: CompanionState;
}

export const PI_OFFICE_CAPABILITY_REGISTRY: CapabilityDefinition[] = [
  {
    id: "office_context",
    label: "Office document context",
    owner: "addin-only",
    tools: ["office_get_context", "office_tool_search", "office_tool_get", "office_batch_execute", "office_read_section", "verify_doc", "get_cell_ranges", "get_all_objects", "search_data", "get_presentation_structure", "get_slide", "list_slide_shapes", "verify_slides"],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "Runs in the Office taskpane through Office.js.",
    group: "works_without_companion",
  },
  {
    id: "office_write",
    label: "Office document edits",
    owner: "addin-only",
    tools: ["office_apply_edit", "edit_doc_text", "set_cell_range", "clear_cell_range", "resize_range", "copy_to", "modify_sheet_structure", "modify_object", "office_navigate", "modify_presentation_structure", "duplicate_slide", "insert_slide_element", "remove_slide_element", "edit_slide_text", "edit_slide_xml", "edit_slide_master", "edit_slide_chart", "copy_image_between_slides", "insert_icon", "office_execute_js"],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "Mutates the active Office document only from the taskpane host.",
    group: "works_without_companion",
  },
  {
    id: "office_review",
    label: "Reviewable edit cards",
    owner: "addin-only",
    tools: ["edit_doc_list", "office_propose_edits", "ask_user"],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "Uses taskpane UI review and permission prompts.",
    group: "works_without_companion",
  },
  {
    id: "office_visual_snapshot",
    label: "Office.js visual/context snapshots",
    owner: "addin-only",
    tools: ["office_capture_snapshot", "verify_doc_visual", "read_range_image", "verify_slide_visual", "get_range_as_csv", "extract_chart_xml", "search_icons"],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "Uses Office.js context, selected-object snapshots, or synthetic metadata; not a full OS screenshot.",
    group: "works_without_companion",
  },
  {
    id: "native_viewport_capture",
    label: "True viewport/window screenshot",
    owner: "companion-only",
    tools: ["office_capture_viewport"],
    requiresCompanion: true,
    requiresSavedDocument: false,
    honestyLabel: "Requires the companion native capture backend; unavailable in taskpane-only mode.",
    group: "requires_companion",
  },
  {
    id: "powerpoint_native_visual",
    label: "PowerPoint native visual verification",
    owner: "addin-only",
    tools: ["verify_slide_visual", "copy_image_between_slides"],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "Uses PowerPoint Office.js slide/shape image paths where supported.",
    group: "works_without_companion",
  },
  {
    id: "local_files",
    label: "Saved document folder reads",
    owner: "companion-only",
    tools: ["read", "grep", "find", "ls"],
    requiresCompanion: true,
    requiresSavedDocument: true,
    honestyLabel: "Reads only from the saved document folder through the companion.",
    group: "requires_companion",
  },
  {
    id: "mcp_connectors",
    label: "Read-safe MCP connectors",
    owner: "companion-only",
    tools: ["mcp", "mcp_tool_search", "mcp_batch_execute"],
    requiresCompanion: true,
    requiresSavedDocument: false,
    honestyLabel: "Runs verified read-safe local stdio or remote HTTP MCP tools through the companion.",
    group: "requires_companion",
  },
  {
    id: "shell_sandbox",
    label: "Sandboxed shell",
    owner: "companion-only",
    tools: ["bash"],
    requiresCompanion: true,
    requiresSavedDocument: true,
    honestyLabel: "Runs only through the companion sandbox after isolation probes pass.",
    group: "requires_companion",
  },
  {
    id: "inference",
    label: "Model inference",
    owner: "either",
    tools: [],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "Taskpane API-key inference remains available; companion inference is preferred only when companion provider auth is available.",
    group: "enhanced_by_companion",
  },
  {
    id: "provider_auth",
    label: "Provider auth and token brokerage",
    owner: "either",
    tools: [],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "Browser API keys stay taskpane-local; OAuth/subscription auth belongs in the companion.",
    group: "enhanced_by_companion",
  },
  {
    id: "image_generation",
    label: "Image generation",
    owner: "either",
    tools: ["generate_image"],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "OpenAI image generation is browser-supported today; future provider breadth should prefer companion auth.",
    group: "enhanced_by_companion",
  },
  {
    id: "diagnostics",
    label: "Runtime diagnostics",
    owner: "either",
    tools: [],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "Browser diagnostics work everywhere; machine runtime checks are companion-enhanced.",
    group: "enhanced_by_companion",
  },
  {
    id: "memory",
    label: "Long-term memory",
    owner: "companion-only",
    tools: [],
    requiresCompanion: true,
    requiresSavedDocument: false,
    honestyLabel: "Durable memory is reserved for companion-owned advanced mode.",
    group: "requires_companion",
  },
  {
    id: "background_jobs",
    label: "Background jobs",
    owner: "either",
    tools: [],
    requiresCompanion: false,
    requiresSavedDocument: false,
    honestyLabel: "Small browser jobs remain taskpane-local; long-running machine jobs should prefer the companion when implemented.",
    group: "enhanced_by_companion",
  },
];

export function getCapabilityDefinition(id: CapabilityId): CapabilityDefinition {
  const definition = PI_OFFICE_CAPABILITY_REGISTRY.find((entry) => entry.id === id);
  if (!definition) {
    throw new Error(`Unknown Pi-Office capability: ${id}`);
  }
  return definition;
}

function unavailableReason(definition: CapabilityDefinition, input: CapabilityResolutionInput): string | undefined {
  if (definition.requiresSavedDocument && !input.documentSaved) {
    return "Save the document before this capability can use saved-document context.";
  }
  if (definition.requiresCompanion && input.companion.status !== "connected") {
    return "The optional companion is not connected.";
  }
  return undefined;
}

function companionCapabilityAvailable(definition: CapabilityDefinition, input: CapabilityResolutionInput): boolean {
  if (input.companion.status !== "connected") return false;
  const capabilities = input.companion.capabilities;
  if (definition.id === "native_viewport_capture") {
    return capabilities.nativeCapture?.state === "available" &&
      capabilities.nativeCapture.hosts.includes(input.host);
  }
  if (definition.id === "local_files") {
    return input.documentSaved && capabilities.fileRead;
  }
  if (definition.id === "mcp_connectors") {
    return Boolean(capabilities.mcp?.available ?? capabilities.localMcp) && Boolean(input.companion.connectorToolNames?.length);
  }
  if (definition.id === "shell_sandbox") {
    return input.documentSaved && capabilities.shell?.state === "available";
  }
  if (definition.id === "inference") {
    return capabilities.agent?.state === "available" && capabilities.providerAuth?.state === "available";
  }
  if (definition.id === "provider_auth") {
    return capabilities.providerAuth?.state === "available";
  }
  if (definition.id === "memory") {
    return capabilities.memory?.state === "available";
  }
  return true;
}

export function resolvePiOfficeCapabilities(input: CapabilityResolutionInput): CapabilityResolution[] {
  return PI_OFFICE_CAPABILITY_REGISTRY.map((definition): CapabilityResolution => {
    if (definition.owner === "addin-only") {
      const hostMismatch =
        definition.id === "powerpoint_native_visual" && input.host !== "powerpoint";
      return {
        ...definition,
        available: !hostMismatch,
        preferredRuntime: "addin",
        activeRuntime: hostMismatch ? undefined : "addin",
        reason: hostMismatch ? "PowerPoint native visual tools are only available in PowerPoint." : undefined,
      };
    }

    if (definition.owner === "companion-only") {
      const reason = unavailableReason(definition, input);
      const available = !reason && companionCapabilityAvailable(definition, input);
      return {
        ...definition,
        available,
        preferredRuntime: "companion",
        activeRuntime: available ? "companion" : undefined,
        reason: available ? undefined : reason ?? "The companion does not currently advertise this capability.",
      };
    }

    const companionAvailable = companionCapabilityAvailable(definition, input);
    return {
      ...definition,
      available: true,
      preferredRuntime: companionAvailable ? "companion" : "addin",
      activeRuntime: companionAvailable ? "companion" : "addin",
      fallbackRuntime: companionAvailable ? "addin" : undefined,
      reason: companionAvailable ? undefined : "Using taskpane runtime until companion support is configured.",
    };
  });
}

export function getResolvedCapability(
  capabilities: readonly CapabilityResolution[],
  id: CapabilityId,
): CapabilityResolution {
  const resolution = capabilities.find((entry) => entry.id === id);
  if (!resolution) {
    throw new Error(`Capability ${id} was not resolved.`);
  }
  return resolution;
}

export function getAvailableToolNames(capabilities: readonly CapabilityResolution[]): Set<string> {
  const names = new Set<string>();
  for (const capability of capabilities) {
    if (!capability.available) continue;
    for (const tool of capability.tools) names.add(tool);
  }
  return names;
}
