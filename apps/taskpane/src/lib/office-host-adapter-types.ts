import type {
  OfficeAnchor,
  OfficeContextPayload,
  OfficeHost,
  OfficeHostAction,
  OfficeStateUpdate,
} from "@pi-office/pi-office-pack/protocol";

export interface OfficeCaptureOptions {
  includeFormatting?: boolean;
  maxImages?: number;
  scope?: string;
}

export interface OfficeHostAdapter {
  host: OfficeHost;
  collectState(base: OfficeStateUpdate): Promise<OfficeStateUpdate>;
  collectContext(base: OfficeStateUpdate, options: OfficeCaptureOptions): Promise<OfficeContextPayload>;
  navigateAnchor(anchor: OfficeAnchor): Promise<unknown>;
  applyAction(action: OfficeHostAction): Promise<unknown>;
}
