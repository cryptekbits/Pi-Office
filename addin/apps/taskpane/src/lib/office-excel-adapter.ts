import type { OfficeAnchor, OfficeContextPayload, OfficeHostAction, OfficeStateUpdate } from "@pi-office/pi-office-pack/protocol";

import type { OfficeCaptureOptions, OfficeHostAdapter } from "./office-host-adapter-types";

interface ExcelOfficeHostAdapterDependencies {
  collectState(base: OfficeStateUpdate): Promise<OfficeStateUpdate>;
  collectContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload>;
  navigateAnchor(anchor: OfficeAnchor): Promise<unknown>;
  applyAction(action: OfficeHostAction): Promise<unknown>;
}

export function createExcelOfficeHostAdapter(dependencies: ExcelOfficeHostAdapterDependencies): OfficeHostAdapter {
  return {
    host: "excel",
    collectState: dependencies.collectState,
    collectContext: dependencies.collectContext,
    navigateAnchor: dependencies.navigateAnchor,
    applyAction: dependencies.applyAction,
  };
}
