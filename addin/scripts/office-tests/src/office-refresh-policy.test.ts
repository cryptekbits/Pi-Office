import assert from "node:assert/strict";
import test from "node:test";

import {
  formatOfficeRefreshError,
  getOfficeRefreshErrorDecision,
  OFFICE_REFRESH_ERROR_DEDUP_WINDOW_MS,
  shouldApplyOfficeRefreshResult,
} from "../../../apps/taskpane/src/lib/office-refresh-policy.js";

test("office refresh policy accepts only the latest active attempt for the same session", () => {
  assert.equal(
    shouldApplyOfficeRefreshResult(
      { sequence: 3, sessionId: "session-a" },
      { active: true, latestSequence: 3, sessionId: "session-a" },
    ),
    true,
  );
  assert.equal(
    shouldApplyOfficeRefreshResult(
      { sequence: 2, sessionId: "session-a" },
      { active: true, latestSequence: 3, sessionId: "session-a" },
    ),
    false,
  );
  assert.equal(
    shouldApplyOfficeRefreshResult(
      { sequence: 3, sessionId: "session-a" },
      { active: true, latestSequence: 3, sessionId: "session-b" },
    ),
    false,
  );
  assert.equal(
    shouldApplyOfficeRefreshResult(
      { sequence: 3, sessionId: "session-a" },
      { active: false, latestSequence: 3, sessionId: "session-a" },
    ),
    false,
  );
});

test("office refresh error policy dedupes repeated transient failures", () => {
  const first = getOfficeRefreshErrorDecision(undefined, "Selection unavailable", 1_000);
  assert.equal(first.shouldSurface, true);
  assert.equal(first.nextRecord.suppressedCount, 0);

  const repeated = getOfficeRefreshErrorDecision(first.nextRecord, "Selection unavailable", 2_000);
  assert.equal(repeated.shouldSurface, false);
  assert.equal(repeated.nextRecord.suppressedCount, 1);
  assert.equal(repeated.nextRecord.lastShownAt, 1_000);

  const afterWindow = getOfficeRefreshErrorDecision(
    repeated.nextRecord,
    "Selection unavailable",
    2_000 + OFFICE_REFRESH_ERROR_DEDUP_WINDOW_MS,
  );
  assert.equal(afterWindow.shouldSurface, true);
  assert.equal(afterWindow.nextRecord.suppressedCount, 0);
  assert.equal(afterWindow.nextRecord.firstSeenAt, 1_000);

  const different = getOfficeRefreshErrorDecision(afterWindow.nextRecord, "Document busy", 3_000);
  assert.equal(different.shouldSurface, true);
  assert.equal(different.nextRecord.message, "Document busy");
  assert.equal(different.nextRecord.firstSeenAt, 3_000);
});

test("office refresh error formatting normalizes thrown values", () => {
  assert.equal(formatOfficeRefreshError(new Error("  Host is busy  ")), "Host is busy");
  assert.equal(formatOfficeRefreshError("  Selection changed  "), "Selection changed");
  assert.equal(formatOfficeRefreshError({ code: "ItemNotFound" }), "{\"code\":\"ItemNotFound\"}");
});
