import type { OfficeAnchor, OfficeContextPayload, OfficeHostAction, OfficeStateUpdate } from "@pi-office/pi-office-pack/protocol";

import type { OfficeCaptureOptions, OfficeHostAdapter } from "./office-host-adapter-types";

interface WordOfficeHostAdapterDependencies {
  collectState(base: OfficeStateUpdate): Promise<OfficeStateUpdate>;
  collectContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload>;
  navigateAnchor(anchor: OfficeAnchor): Promise<unknown>;
  applyAction(action: OfficeHostAction): Promise<unknown>;
}

export function createWordOfficeHostAdapter(dependencies: WordOfficeHostAdapterDependencies): OfficeHostAdapter {
  return {
    host: "word",
    collectState: dependencies.collectState,
    collectContext: dependencies.collectContext,
    navigateAnchor: dependencies.navigateAnchor,
    applyAction: dependencies.applyAction,
  };
}
