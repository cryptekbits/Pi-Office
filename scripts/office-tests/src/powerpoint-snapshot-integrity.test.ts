import assert from "node:assert/strict";
import test from "node:test";

import { captureDocumentSnapshot } from "../../../apps/taskpane/src/lib/office/document-tools.js";

type SlicePlan = {
  status: "succeeded" | "failed";
  data?: unknown;
  error?: string;
};

type PowerPointSliceResult =
  | {
    status: string;
    value: {
      data: unknown;
    };
  }
  | {
    status: string;
    error: {
      message: string;
    };
  };

function installPowerPointSnapshotMock(slicePlans: SlicePlan[]): {
  getCloseCallCount: () => number;
  restore: () => void;
} {
  const globalAny = globalThis as typeof globalThis & {
    Office?: typeof Office;
    btoa?: (value: string) => string;
  };
  const previousOffice = globalAny.Office;
  const previousBtoa = globalAny.btoa;

  if (!globalAny.btoa) {
    globalAny.btoa = (value: string) => Buffer.from(value, "binary").toString("base64");
  }

  let closeCalls = 0;
  const statusSucceeded = "succeeded";
  const statusFailed = "failed";

  const file = {
    sliceCount: slicePlans.length,
    getSliceAsync: (index: number, callback: (result: PowerPointSliceResult) => void) => {
      const plan = slicePlans[index];
      if (!plan) {
        throw new Error(`Missing slice plan for index ${index}`);
      }
      if (plan.status === "succeeded") {
        callback({
          status: statusSucceeded,
          value: { data: plan.data },
        });
        return;
      }
      callback({
        status: statusFailed,
        error: { message: plan.error ?? `Slice ${index} failed` },
      });
    },
    closeAsync: (callback?: () => void) => {
      closeCalls += 1;
      callback?.();
    },
  };

  globalAny.Office = {
    FileType: { Compressed: "compressed" },
    AsyncResultStatus: {
      Succeeded: statusSucceeded,
      Failed: statusFailed,
    },
    context: {
      document: {
        getFileAsync: (
          _fileType: unknown,
          _options: unknown,
          callback: (result: { status: string; value: typeof file }) => void,
        ) => {
          callback({
            status: statusSucceeded,
            value: file,
          });
        },
      },
    },
  } as unknown as typeof Office;

  return {
    getCloseCallCount: () => closeCalls,
    restore: () => {
      globalAny.Office = previousOffice;
      if (previousBtoa === undefined) {
        delete globalAny.btoa;
      } else {
        globalAny.btoa = previousBtoa;
      }
    },
  };
}

test("captureDocumentSnapshot rejects PowerPoint snapshot when any slice retrieval fails", async () => {
  const mock = installPowerPointSnapshotMock([
    { status: "succeeded", data: "chunk-0" },
    { status: "failed", error: "slice retrieval failed" },
  ]);

  try {
    await assert.rejects(
      captureDocumentSnapshot("powerpoint"),
      /slice/i,
    );
    assert.equal(mock.getCloseCallCount(), 1);
  } finally {
    mock.restore();
  }
});

test("captureDocumentSnapshot rejects PowerPoint snapshot when any chunk data is missing", async () => {
  const mock = installPowerPointSnapshotMock([
    { status: "succeeded", data: "chunk-0" },
    { status: "succeeded", data: undefined },
  ]);

  try {
    await assert.rejects(
      captureDocumentSnapshot("powerpoint"),
      /missing/i,
    );
    assert.equal(mock.getCloseCallCount(), 1);
  } finally {
    mock.restore();
  }
});

test("captureDocumentSnapshot rejects PowerPoint snapshot when any chunk payload is corrupt", async () => {
  const mock = installPowerPointSnapshotMock([
    { status: "succeeded", data: "chunk-0" },
    { status: "succeeded", data: ["not-a-byte"] },
  ]);

  try {
    await assert.rejects(
      captureDocumentSnapshot("powerpoint"),
      /missing/i,
    );
    assert.equal(mock.getCloseCallCount(), 1);
  } finally {
    mock.restore();
  }
});
