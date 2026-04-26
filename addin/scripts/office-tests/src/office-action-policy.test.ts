import assert from "node:assert/strict";
import test from "node:test";

import { assertDestructiveActionAllowed, isDestructiveOfficeAction } from "../../../apps/taskpane/src/lib/office-action-policy.js";

test("isDestructiveOfficeAction identifies Excel and PowerPoint destructive actions", () => {
  assert.deepEqual(isDestructiveOfficeAction({ type: "deleteRows" }), {
    destructive: true,
    description: "delete worksheet rows",
  });
  assert.deepEqual(isDestructiveOfficeAction({ type: "deleteSlide" }), {
    destructive: true,
    description: "delete slides from the presentation",
  });
  assert.deepEqual(isDestructiveOfficeAction({ type: "addSlideChart" }), {
    destructive: false,
  });
});

test("assertDestructiveActionAllowed blocks destructive actions without explicit confirmation", () => {
  assert.throws(
    () => assertDestructiveActionAllowed("powerpoint", { type: "deleteSlide", target: { kind: "slide", slideId: "slide-2" } }),
    (error: unknown) => {
      assert.ok(error instanceof Error);
      assert.equal((error as Error & { code?: string }).code, "destructive_confirmation_required");
      assert.match(error.message, /confirmDestructive=true/);
      return true;
    },
  );
});

test("assertDestructiveActionAllowed allows confirmed destructive actions and leaves safe actions alone", () => {
  assert.doesNotThrow(() =>
    assertDestructiveActionAllowed("excel", {
      type: "deleteRows",
      target: { kind: "range", sheetName: "Budget", address: "5:5" },
      options: { confirmDestructive: true },
    }),
  );

  assert.doesNotThrow(() =>
    assertDestructiveActionAllowed("word", {
      type: "insertText",
      content: "Safe edit",
    }),
  );
});
