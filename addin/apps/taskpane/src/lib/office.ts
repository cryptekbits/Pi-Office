import type {
  OfficeAnchor,
  OfficeContextPayload,
  OfficeHost,
  OfficeHostAction,
  OfficeStateUpdate,
  OfficeVisualSnapshot,
} from "@pi-office/pi-office-pack/protocol";
import { assertDestructiveActionAllowed } from "./office-action-policy";
import type { OfficeCaptureOptions, OfficeHostAdapter } from "./office-host-adapter-types";
import { createExcelOfficeHostAdapter } from "./office-excel-adapter";
import { createPowerPointOfficeHostAdapter } from "./office-powerpoint-adapter";
import { createWordOfficeHostAdapter } from "./office-word-adapter";
import {
  buildBaseState,
  buildOfficeActionResult,
  mapHost,
  toValueMatrix,
} from "./office/shared";
import {
  applyWordAction,
  collectWordContext,
  collectWordState,
  navigateWordAnchor,
  searchWordDocument,
} from "./office/word";
import {
  applyExcelAction,
  collectExcelContext,
  collectExcelState,
  navigateExcelAnchor,
} from "./office/excel";
import {
  applyPowerPointAction,
  collectPowerPointContext,
  collectPowerPointState,
  navigatePowerPointAnchor,
} from "./office/powerpoint";

const officeHostAdapters: Record<OfficeHost, OfficeHostAdapter> = {
  word: createWordOfficeHostAdapter({
    collectState: collectWordState,
    collectContext: collectWordContext,
    navigateAnchor: navigateWordAnchor,
    applyAction: applyWordAction,
  }),
  excel: createExcelOfficeHostAdapter({
    collectState: collectExcelState,
    collectContext: collectExcelContext,
    navigateAnchor: navigateExcelAnchor,
    applyAction: applyExcelAction,
  }),
  powerpoint: createPowerPointOfficeHostAdapter({
    collectState: collectPowerPointState,
    collectContext: collectPowerPointContext,
    navigateAnchor: navigatePowerPointAnchor,
    applyAction: applyPowerPointAction,
  }),
};

function getOfficeHostAdapter(host: OfficeHost): OfficeHostAdapter {
  return officeHostAdapters[host];
}

export async function collectOfficeState(explicitHost?: OfficeHost): Promise<OfficeStateUpdate> {
  const host = explicitHost ?? mapHost(Office.context.host);
  const base = await buildBaseState(host);
  return getOfficeHostAdapter(host).collectState(base);
}

export async function collectOfficeContext(explicitHost?: OfficeHost, options: OfficeCaptureOptions = {}): Promise<OfficeContextPayload> {
  const host = explicitHost ?? mapHost(Office.context.host);
  const base = await buildBaseState(host);
  return getOfficeHostAdapter(host).collectContext(base, options);
}

export async function capturePromptVisuals(explicitHost?: OfficeHost, maxImages = 2): Promise<OfficeVisualSnapshot[]> {
  const payload = await collectOfficeContext(explicitHost, {
    includeFormatting: false,
    maxImages,
  });
  return payload.visuals ?? [];
}

export async function navigateOfficeAnchor(host: OfficeHost, anchor: OfficeAnchor): Promise<unknown> {
  const adapter = getOfficeHostAdapter(host);
  return buildOfficeActionResult(host, "navigate:" + anchor.kind, await adapter.navigateAnchor(anchor), {
    navigation: anchor,
    target: anchor,
  });
}

export async function applyHostAction(host: OfficeHost, action: OfficeHostAction): Promise<unknown> {
  assertDestructiveActionAllowed(host, action);
  const adapter = getOfficeHostAdapter(host);
  return buildOfficeActionResult(host, action.type, await adapter.applyAction(action), { target: action.target });
}

export async function applyHostEdit(host: OfficeHost, params: { mode: string; content: string; format?: string }): Promise<unknown> {
  const action: OfficeHostAction =
    params.format === "matrix" || params.mode === "setRangeValues"
      ? { type: "setRangeValues", values: toValueMatrix(params.content, params.content), format: params.format, content: params.content }
      : {
          type: params.format === "html" ? "insertHtml" : "insertText",
          placement: params.mode === "insertAfterSelection" ? "after" : "replace",
          format: params.format,
          content: params.content,
        };

  return applyHostAction(host, action);
}

export {
  BROWSER_DEBUG_OFFICE_CAPABILITY,
  waitForOfficeReady,
  buildOpenRequest,
  buildConnectorScopeContext,
  createBrowserDebugOfficeState,
  isBrowserDebugOfficeState,
  shouldUseBrowserDebugOfficeState,
  subscribeToOfficeChanges,
  readOfficeTheme,
} from "./office/shared";
export type { OfficeThemeSnapshot } from "./office/shared";

export {
  captureDocumentSnapshot,
  restoreDocumentSnapshot,
  readDocumentSection,
  executeOfficeJs,
  applyAcceptedEdits,
  proposeDocumentEdits,
} from "./office/document-tools";
export { searchWordDocument } from "./office/word";
export type { DocumentSnapshotData } from "./office/document-tools";
