import assert from "node:assert/strict";
import test from "node:test";

import {
  buildExcelRestoreContent,
  restoreDocumentSnapshot,
  type DocumentSnapshotData,
} from "../../../apps/taskpane/src/lib/office/document-tools.js";

type AssignedExcelRange = {
  address?: string | undefined;
  values?: unknown[][] | undefined;
  formulas?: unknown[][] | undefined;
  numberFormat?: string[][] | undefined;
};

type ExcelRunContextMock = {
  workbook: {
    worksheets: {
      getItem: (name: string) => {
        getRange: (address: string) => {
          values: unknown[][];
          formulas: unknown[][];
          numberFormat: string[][];
        };
      };
    };
  };
  sync: () => Promise<void>;
};

function installExcelRestoreMock(): {
  assigned: AssignedExcelRange;
  getSyncCount: () => number;
  restore: () => void;
} {
  const globalAny = globalThis as typeof globalThis & {
    Excel?: typeof Excel;
  };
  const previousExcel = globalAny.Excel;
  const assigned: AssignedExcelRange = {};
  let syncCount = 0;

  const range = {
    set values(value: unknown[][]) {
      assigned.values = value;
    },
    set formulas(value: unknown[][]) {
      assigned.formulas = value;
    },
    set numberFormat(value: string[][]) {
      assigned.numberFormat = value;
    },
  };

  const worksheet = {
    getRange: (address: string) => {
      assigned.address = address;
      return range;
    },
  };

  const context: ExcelRunContextMock = {
    workbook: {
      worksheets: {
        getItem: (name: string) => {
          if (name !== "Model") {
            throw new Error(`Unexpected worksheet: ${name}`);
          }
          return worksheet;
        },
      },
    },
    sync: async () => {
      syncCount += 1;
    },
  };

  globalAny.Excel = {
    run: async <T>(callback: (context: ExcelRunContextMock) => Promise<T>) => callback(context),
  } as unknown as typeof Excel;

  return {
    assigned,
    getSyncCount: () => syncCount,
    restore: () => {
      globalAny.Excel = previousExcel;
    },
  };
}

test("Excel restore content prefers captured formulas and keeps constants in the same matrix", () => {
  const content = buildExcelRestoreContent({
    name: "Model",
    usedRangeAddress: "Model!A1:B2",
    values: [
      ["Revenue", "Growth"],
      [100, 110],
    ],
    formulas: [
      ["Revenue", "Growth"],
      [100, "=A2*1.1"],
    ],
    numberFormats: [
      ["@", "@"],
      ["$#,##0", "$#,##0"],
    ],
  });

  assert.equal(content.property, "formulas");
  assert.deepEqual(content.matrix, [
    ["Revenue", "Growth"],
    [100, "=A2*1.1"],
  ]);
  assert.match(content.warnings.join(" "), /used-range formulas/);
});

test("restoreDocumentSnapshot writes Excel formulas instead of flattening formulas to values", async () => {
  const mock = installExcelRestoreMock();
  const snapshot: DocumentSnapshotData = {
    sheets: [
      {
        name: "Model",
        usedRangeAddress: "Model!A1:B2",
        values: [
          ["Revenue", "Growth"],
          [100, 110],
        ],
        formulas: [
          ["Revenue", "Growth"],
          [100, "=A2*1.1"],
        ],
        numberFormats: [
          ["@", "@"],
          ["$#,##0", "$#,##0"],
        ],
      },
    ],
  };

  try {
    const result = await restoreDocumentSnapshot("excel", snapshot);
    assert.equal(mock.assigned.address, "A1:B2");
    assert.deepEqual(mock.assigned.formulas, [
      ["Revenue", "Growth"],
      [100, "=A2*1.1"],
    ]);
    assert.equal(mock.assigned.values, undefined);
    assert.deepEqual(mock.assigned.numberFormat, [
      ["@", "@"],
      ["$#,##0", "$#,##0"],
    ]);
    assert.equal(mock.getSyncCount(), 1);
    assert.match(result.warnings.join(" "), /tables, charts, data validation/);
  } finally {
    mock.restore();
  }
});

test("restoreDocumentSnapshot falls back to values with a warning when formula shape is unavailable", async () => {
  const mock = installExcelRestoreMock();
  const snapshot: DocumentSnapshotData = {
    sheets: [
      {
        name: "Model",
        usedRangeAddress: "Model!A1:B2",
        values: [
          ["Revenue", "Growth"],
          [100, 110],
        ],
        formulas: [["=A1"]],
        numberFormats: [
          ["@", "@"],
          ["$#,##0", "$#,##0"],
        ],
      },
    ],
  };

  try {
    const result = await restoreDocumentSnapshot("excel", snapshot);
    assert.deepEqual(mock.assigned.values, [
      ["Revenue", "Growth"],
      [100, 110],
    ]);
    assert.equal(mock.assigned.formulas, undefined);
    assert.match(result.warnings.join(" "), /values only/);
  } finally {
    mock.restore();
  }
});
