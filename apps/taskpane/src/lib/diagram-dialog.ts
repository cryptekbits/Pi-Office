const STORAGE_KEY = "pi-expand-viewer-data";
const UPDATE_STORAGE_KEY = "pi-expand-viewer-update";
const VIEWER_PATH = "/diagram-viewer.html";

export type ExpandContentType = "drawio" | "svg" | "image";

export interface ExpandViewerOptions {
  isDark?: boolean | undefined;
  onInsertRequest?: ((pngBase64: string) => void) | undefined;
  onDataUpdate?: ((newData: string) => void) | undefined;
}

let activeDialog: Office.Dialog | null = null;

export async function openExpandViewer(
  contentType: ExpandContentType,
  data: string,
  options?: ExpandViewerOptions,
): Promise<void> {
  const isDark = options?.isDark ?? false;

  localStorage.setItem(
    STORAGE_KEY,
    JSON.stringify({ contentType, data, isDark }),
  );

  const origin = window.location.origin;
  const url = `${origin}${VIEWER_PATH}`;

  if (activeDialog) {
    try { activeDialog.close(); } catch { /* ignore */ }
    activeDialog = null;
  }

  if (hasOfficeDialogApi()) {
    return openOfficeDialog(Office.context.ui, url, options);
  }

  openFallbackWindow(url, options);
}

function hasOfficeDialogApi(): boolean {
  try {
    return (
      typeof Office !== "undefined" &&
      Office.context?.ui?.displayDialogAsync != null
    );
  } catch { return false; }
}

function openOfficeDialog(
  ui: typeof Office.context.ui,
  url: string,
  options?: ExpandViewerOptions,
): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    ui.displayDialogAsync(
      url,
      { width: 80, height: 80, displayInIframe: false },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          reject(new Error(result.error?.message ?? "Failed to open dialog"));
          return;
        }

        const dialog = result.value;
        activeDialog = dialog;

        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          (arg) => {
            const msg = "message" in arg ? String(arg.message ?? "") : "";
            if (msg === "close") {
              dialog.close();
              activeDialog = null;
              const updatedData = localStorage.getItem(UPDATE_STORAGE_KEY);
              if (updatedData) {
                localStorage.removeItem(UPDATE_STORAGE_KEY);
                options?.onDataUpdate?.(updatedData);
              }
              resolve();
            } else if (msg.startsWith("insert:") && options?.onInsertRequest) {
              options.onInsertRequest(msg.slice(7));
            }
          },
        );

        dialog.addEventHandler(
          Office.EventType.DialogEventReceived,
          () => {
            activeDialog = null;
            const updatedData = localStorage.getItem(UPDATE_STORAGE_KEY);
            if (updatedData) {
              localStorage.removeItem(UPDATE_STORAGE_KEY);
              options?.onDataUpdate?.(updatedData);
            }
            resolve();
          },
        );
      },
    );
  });
}

function openFallbackWindow(
  url: string,
  options?: ExpandViewerOptions,
): void {
  const w = Math.round(screen.width * 0.8);
  const h = Math.round(screen.height * 0.8);
  const left = Math.round((screen.width - w) / 2);
  const top = Math.round((screen.height - h) / 2);
  const win = window.open(
    url,
    "pi-expand-viewer",
    `width=${w},height=${h},left=${left},top=${top},resizable=yes,scrollbars=no`,
  );

  if (!win) return;

  const handler = (event: MessageEvent) => {
    if (event.data?.type !== "pi-expand-viewer") return;
    const msg = String(event.data.payload ?? "");
    if (msg === "close") {
      win.close();
      window.removeEventListener("message", handler);
      const updatedData = localStorage.getItem(UPDATE_STORAGE_KEY);
      if (updatedData) {
        localStorage.removeItem(UPDATE_STORAGE_KEY);
        options?.onDataUpdate?.(updatedData);
      }
    } else if (msg.startsWith("insert:") && options?.onInsertRequest) {
      options.onInsertRequest(msg.slice(7));
    }
  };

  window.addEventListener("message", handler);
}
