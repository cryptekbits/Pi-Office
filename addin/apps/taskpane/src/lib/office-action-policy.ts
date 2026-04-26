import type { OfficeHost, OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import { HOST_LABELS } from "@pi-office/pi-office-pack/defaults";

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === "object" && !Array.isArray(value));
}

function trimString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

function toBoolean(value: unknown): boolean | undefined {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    if (value === "true") return true;
    if (value === "false") return false;
  }

  return undefined;
}

function getActionOptions(action: OfficeHostAction): Record<string, unknown> {
  return isRecord(action.options) ? action.options : {};
}

export function createOfficeActionError(
  code: string,
  message: string,
  details: Record<string, unknown> = {},
): Error & { code: string; details?: Record<string, unknown> } {
  const error = new Error(message) as Error & { code: string; details?: Record<string, unknown> };
  error.code = code;
  if (Object.keys(details).length) {
    error.details = details;
  }
  return error;
}

export function isDestructiveOfficeAction(action: OfficeHostAction): { destructive: boolean; description?: string | undefined } {
  const type = trimString(action.type)?.toLowerCase() ?? "";

  switch (type) {
    case "deletecomment":
      return { destructive: true, description: "delete the target Word comment" };
    case "acceptrevision":
      return { destructive: true, description: "accept the targeted tracked change" };
    case "rejectrevision":
      return { destructive: true, description: "reject the targeted tracked change" };
    case "acceptallrevisions":
      return { destructive: true, description: "accept all tracked changes in scope" };
    case "rejectallrevisions":
      return { destructive: true, description: "reject all tracked changes in scope" };
    case "deleterows":
      return { destructive: true, description: "delete worksheet rows" };
    case "deletecolumns":
      return { destructive: true, description: "delete worksheet columns" };
    case "removeduplicates":
      return { destructive: true, description: "remove duplicate rows from the selected range" };
    case "deleteworksheet":
      return { destructive: true, description: "delete the worksheet" };
    case "cleardatavalidation":
      return { destructive: true, description: "clear data-validation rules from the selected range" };
    case "clearconditionalformats":
      return { destructive: true, description: "clear conditional formatting from the selected range" };
    case "clearshapetext":
      return { destructive: true, description: "clear all text from the selected shape" };
    case "deleteshape":
    case "deleteshapes":
      return { destructive: true, description: "delete shapes from the slide" };
    case "deleteslide":
    case "deleteslides":
      return { destructive: true, description: "delete slides from the presentation" };
    case "deletetablerows":
      return { destructive: true, description: "delete table rows from the selected PowerPoint table" };
    case "deletetablecolumns":
      return { destructive: true, description: "delete table columns from the selected PowerPoint table" };
    case "cleartable":
      return { destructive: true, description: "clear all cell content from the selected PowerPoint table" };
    default:
      return { destructive: false };
  }
}

export function assertDestructiveActionAllowed(host: OfficeHost, action: OfficeHostAction): void {
  const assessment = isDestructiveOfficeAction(action);
  if (!assessment.destructive) {
    return;
  }

  const options = getActionOptions(action);
  const confirmed = [options.confirmDestructive, action.confirmDestructive, options.allowDestructive, action.allowDestructive, options.force, action.force]
    .map((value) => toBoolean(value))
    .some((value) => value === true);
  if (confirmed) {
    return;
  }

  const description = assessment.description ?? `${HOST_LABELS[host]} ${action.type}`;
  throw createOfficeActionError(
    "destructive_confirmation_required",
    `${HOST_LABELS[host]} ${action.type} is blocked because it would ${description}. Re-run with action.options.confirmDestructive=true after verifying the target.`,
    {
      host,
      actionType: action.type,
      target: action.target,
      confirmationKey: "action.options.confirmDestructive",
    },
  );
}
