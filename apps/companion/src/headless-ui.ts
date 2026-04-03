import type { ExtensionUIContext } from "@mariozechner/pi-coding-agent";

export const headlessUiContext: ExtensionUIContext = {
  select: async () => undefined,
  confirm: async () => false,
  input: async () => undefined,
  notify: () => {},
  onTerminalInput: () => () => {},
  setStatus: () => {},
  setWorkingMessage: () => {},
  setHiddenThinkingLabel: () => {},
  setWidget: () => {},
  setFooter: () => {},
  setHeader: () => {},
  setTitle: () => {},
  custom: async () => {
    throw new Error("Custom extension UI is unavailable in the Office companion.");
  },
  pasteToEditor: () => {},
  setEditorText: () => {},
  getEditorText: () => "",
  editor: async () => undefined,
  setEditorComponent: () => {},
  theme: {} as ExtensionUIContext["theme"],
  getAllThemes: () => [],
  getTheme: () => undefined,
  setTheme: () => ({ success: false, error: "Themes are unavailable in headless mode." }),
  getToolsExpanded: () => true,
  setToolsExpanded: () => {},
};
