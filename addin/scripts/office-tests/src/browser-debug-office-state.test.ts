import assert from "node:assert/strict";
import test from "node:test";

import {
  BROWSER_DEBUG_OFFICE_CAPABILITY,
  createBrowserDebugOfficeState,
  isBrowserDebugOfficeState,
  shouldUseBrowserDebugOfficeState,
} from "../../../apps/taskpane/src/lib/office/shared.js";

test("browser debug Office state defaults to an honest unsaved Word preview", () => {
  const state = createBrowserDebugOfficeState("");

  assert.equal(state.host, "word");
  assert.equal(state.document.title, "Browser Preview (Word)");
  assert.equal(state.document.saved, false);
  assert.equal(state.selection.label, "No Office host attached");
  assert.equal(state.capabilities.includes(BROWSER_DEBUG_OFFICE_CAPABILITY), true);
  assert.equal(isBrowserDebugOfficeState(state), true);
});

test("browser debug Office state can target other host chrome by query string", () => {
  assert.equal(createBrowserDebugOfficeState("?piOfficeHost=excel").host, "excel");
  assert.equal(createBrowserDebugOfficeState("?piOfficeHost=powerpoint").host, "powerpoint");
  assert.equal(createBrowserDebugOfficeState("?piOfficeHost=unknown").host, "word");
});

test("browser debug fallback is dev-only unless explicitly enabled", () => {
  const error = new Error("Office host is not available. Open the add-in inside Word, Excel, or PowerPoint.");

  assert.equal(shouldUseBrowserDebugOfficeState(error, { isDev: true }), true);
  assert.equal(shouldUseBrowserDebugOfficeState(error, { isDev: false }), false);
  assert.equal(shouldUseBrowserDebugOfficeState(error, { isDev: false, search: "?piOfficeBrowserDebug=1" }), true);
  assert.equal(shouldUseBrowserDebugOfficeState(error, { isDev: true, search: "?officeDebug=0" }), false);
  assert.equal(shouldUseBrowserDebugOfficeState(new Error("Provider failed"), { isDev: true }), false);
});
