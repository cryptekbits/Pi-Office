import type { OfficeToolRequest, OfficeToolResult } from "@pi-office/pi-office-pack/protocol";
import { createOfficeToolExecutor } from "./office-bridge";
import {
  applyHostAction,
  collectOfficeContext,
  executeOfficeJs,
  navigateOfficeAnchor,
  proposeDocumentEdits,
  readDocumentSection,
} from "./office";

export const executeOfficeTool = createOfficeToolExecutor({
  collectOfficeContext,
  applyHostAction,
  navigateOfficeAnchor,
  readDocumentSection,
  executeOfficeJs,
  proposeEdits: proposeDocumentEdits,
  logger: console,
});
