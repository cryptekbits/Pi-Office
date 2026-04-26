import type { OfficeAnchor, OfficeContextPayload, OfficeHostAction, OfficeStateUpdate } from "@pi-office/pi-office-pack/protocol";

import type { OfficeCaptureOptions, OfficeHostAdapter } from "./office-host-adapter-types";

interface PowerPointOfficeHostAdapterDependencies {
  collectState(base: OfficeStateUpdate): Promise<OfficeStateUpdate>;
  collectContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload>;
  navigateAnchor(anchor: OfficeAnchor): Promise<unknown>;
  applyAction(action: OfficeHostAction): Promise<unknown>;
}

export function createPowerPointOfficeHostAdapter(dependencies: PowerPointOfficeHostAdapterDependencies): OfficeHostAdapter {
  return {
    host: "powerpoint",
    collectState: dependencies.collectState,
    collectContext: dependencies.collectContext,
    navigateAnchor: dependencies.navigateAnchor,
    applyAction: dependencies.applyAction,
  };
}
