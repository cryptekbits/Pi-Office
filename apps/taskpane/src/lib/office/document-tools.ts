import type { OfficeHost } from "@pi-office/pi-office-pack/protocol";
import {
  supportsRequirementSet,
} from "./shared";

export interface DocumentSnapshotData {
  ooxml?: string;
  sheets?: Array<{
    name: string;
    usedRangeAddress: string;
    values: unknown[][];
    numberFormats: string[][];
    formulas: string[][];
  }>;
  presentationBase64?: string;
}

export async function captureDocumentSnapshot(host: OfficeHost): Promise<DocumentSnapshotData> {
  if (host === "word") {
    return Word.run(async (context) => {
      const body = context.document.body;
      const ooxml = body.getOoxml();
      await context.sync();
      return { ooxml: ooxml.value };
    });
  }

  if (host === "excel") {
    return Excel.run(async (context) => {
      const sheets = context.workbook.worksheets;
      sheets.load("items/name");
      await context.sync();
      const sheetData: DocumentSnapshotData["sheets"] = [];
      for (const sheet of sheets.items) {
        const usedRange = sheet.getUsedRangeOrNullObject(true);
        usedRange.load("address,values,numberFormat,formulas");
        await context.sync();
        if (usedRange.isNullObject) continue;
        sheetData.push({
          name: sheet.name,
          usedRangeAddress: usedRange.address,
          values: usedRange.values as unknown[][],
          numberFormats: usedRange.numberFormat as string[][],
          formulas: usedRange.formulas as string[][],
        });
      }
      return { sheets: sheetData };
    });
  }

  if (host === "powerpoint") {
    return new Promise<DocumentSnapshotData>((resolve, reject) => {
      Office.context.document.getFileAsync(
        Office.FileType.Compressed,
        { sliceSize: 4194304 },
        (result) => {
          if (result.status !== Office.AsyncResultStatus.Succeeded) {
            reject(new Error("Failed to get PowerPoint file"));
            return;
          }
          const file = result.value;
          const sliceCount = file.sliceCount;
          const chunks: string[] = [];
          let received = 0;
          for (let i = 0; i < sliceCount; i++) {
            file.getSliceAsync(i, (sliceResult) => {
              if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                const raw = sliceResult.value.data;
                const bytes = typeof raw === "string" ? raw : btoa(String.fromCharCode(...new Uint8Array(raw)));
                chunks[i] = bytes;
              }
              received++;
              if (received === sliceCount) {
                file.closeAsync();
                resolve({ presentationBase64: chunks.join("") });
              }
            });
          }
        },
      );
    });
  }

  return {};
}

export async function restoreDocumentSnapshot(host: OfficeHost, data: DocumentSnapshotData): Promise<void> {
  if (host === "word" && data.ooxml) {
    return Word.run(async (context) => {
      context.document.body.insertOoxml(data.ooxml!, "Replace");
      await context.sync();
    });
  }

  if (host === "excel" && data.sheets?.length) {
    return Excel.run(async (context) => {
      for (const sheetData of data.sheets!) {
        let sheet: Excel.Worksheet;
        try {
          sheet = context.workbook.worksheets.getItem(sheetData.name);
        } catch {
          continue;
        }
        const addr = sheetData.usedRangeAddress.includes("!")
          ? sheetData.usedRangeAddress.split("!")[1]
          : sheetData.usedRangeAddress;
        if (!addr) continue;
        const range = sheet.getRange(addr);
        range.values = sheetData.values as (string | number | boolean)[][];
        range.numberFormat = sheetData.numberFormats as string[][];
        await context.sync();
      }
    });
  }

  if (host === "powerpoint" && data.presentationBase64) {
    return PowerPoint.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load("items/id");
      await context.sync();
      for (const slide of slides.items) {
        slide.delete();
      }
      await context.sync();
      context.presentation.insertSlidesFromBase64(data.presentationBase64!, {
        formatting: "UseDestinationTheme" as PowerPoint.InsertSlideFormatting,
      });
      await context.sync();
    });
  }
}

// ---------------------------------------------------------------------------
// read_doc_section — paginated paragraph reading
// ---------------------------------------------------------------------------

export async function readDocumentSection(
  host: OfficeHost,
  startIndex: number,
  endIndex: number,
  includeStyles: boolean,
): Promise<unknown> {
  if (host !== "word") {
    return { error: "office_read_section is only supported for Word documents." };
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    const supportsParagraphIds = supportsRequirementSet("WordApi", "1.6");
    paragraphs.load(
      includeStyles
        ? supportsParagraphIds
          ? "items/text,items/uniqueLocalId,items/style,items/styleBuiltIn"
          : "items/text,items/style,items/styleBuiltIn"
        : supportsParagraphIds
          ? "items/text,items/uniqueLocalId"
          : "items/text",
    );
    await context.sync();

    const total = paragraphs.items.length;
    const safeStart = Math.max(0, Math.min(startIndex, total));
    const safeEnd = Math.max(safeStart, Math.min(endIndex, total));
    const slice = paragraphs.items.slice(safeStart, safeEnd);

    const result = slice.map((p, i) => {
      const entry: Record<string, unknown> = {
        index: safeStart + i,
        text: p.text,
      };
      if (supportsParagraphIds) entry.paragraphId = p.uniqueLocalId;
      if (includeStyles) {
        entry.style = p.styleBuiltIn || p.style;
        entry.isHeading = /heading/i.test(String(p.styleBuiltIn || p.style || ""));
      }
      return entry;
    });

    return {
      paragraphs: result,
      startIndex: safeStart,
      endIndex: safeEnd,
      totalParagraphs: total,
      hasMore: safeEnd < total,
    };
  });
}

// ---------------------------------------------------------------------------
// execute_office_js — arbitrary Office.js code execution
// ---------------------------------------------------------------------------

const BLOCKED_JS_PATTERNS = [
  /\bfetch\s*\(/i,
  /\bXMLHttpRequest\b/i,
  /\blocalStorage\b/i,
  /\bsessionStorage\b/i,
  /\bindexedDB\b/i,
  /\bWebSocket\b/i,
  /\beval\s*\(/i,
  /\bFunction\s*\(/i,
  /\bimport\s*\(/i,
  /\bnavigator\b/i,
  /\bdocument\.cookie\b/i,
];

export async function executeOfficeJs(_host: OfficeHost, code: string): Promise<unknown> {
  for (const pattern of BLOCKED_JS_PATTERNS) {
    if (pattern.test(code)) {
      return {
        error: `Code blocked: contains disallowed pattern "${pattern.source}". office_execute_js must not access network, storage, or eval.`,
      };
    }
  }

  try {
    const asyncFn = new Function("Word", "Excel", "PowerPoint", "Office", `
      "use strict";
      return (async () => {
        ${code}
      })();
    `);
    const result = await asyncFn(
      typeof Word !== "undefined" ? Word : undefined,
      typeof Excel !== "undefined" ? Excel : undefined,
      typeof PowerPoint !== "undefined" ? PowerPoint : undefined,
      typeof Office !== "undefined" ? Office : undefined,
    );
    return { ok: true, result: result ?? null };
  } catch (error) {
    return {
      ok: false,
      error: error instanceof Error ? error.message : String(error),
      stack: error instanceof Error ? error.stack : undefined,
    };
  }
}

// ---------------------------------------------------------------------------
// propose_edits — batch edit proposals for user review
// ---------------------------------------------------------------------------

export async function applyAcceptedEdits(
  edits: Array<{
    searchText: string;
    newText: string;
    kind: string;
  }>,
): Promise<{ applied: number; failed: number; errors: string[] }> {
  if (!edits.length) return { applied: 0, failed: 0, errors: [] };

  return Word.run(async (context) => {
    let applied = 0;
    let failed = 0;
    const errors: string[] = [];

    for (const edit of edits) {
      try {
        const searchText = edit.searchText;
        if (!searchText) {
          errors.push(`Edit missing searchText, skipped.`);
          failed++;
          continue;
        }

        if (searchText.length <= 255) {
          const results = context.document.body.search(searchText, { matchCase: true, matchWholeWord: false });
          results.load("items");
          await context.sync();

          if (!results.items.length) {
            errors.push(`Text not found: "${searchText.slice(0, 60)}..."`);
            failed++;
            continue;
          }

          const target = results.items[0]!;
          if (edit.kind === "delete") {
            target.delete();
          } else {
            target.insertText(edit.newText, "Replace");
          }
          await context.sync();
          applied++;
        } else {
          // Fallback for long searchText: search a short prefix, then verify full match via paragraph text
          const prefix = searchText.slice(0, 200);
          const results = context.document.body.search(prefix, { matchCase: true, matchWholeWord: false });
          results.load("items/text,items/paragraphs/items/text");
          await context.sync();

          let matched = false;
          for (const candidate of results.items) {
            const para = candidate.paragraphs.getFirst();
            para.load("text");
            await context.sync();
            if (para.text.includes(searchText)) {
              const fullRange = para.search(searchText.slice(0, 255), { matchCase: true });
              fullRange.load("items");
              await context.sync();
              if (fullRange.items.length) {
                if (edit.kind === "delete") {
                  fullRange.items[0]!.delete();
                } else {
                  fullRange.items[0]!.insertText(edit.newText, "Replace");
                }
                await context.sync();
                applied++;
                matched = true;
                break;
              }
            }
          }
          if (!matched) {
            errors.push(`Long text not found: "${searchText.slice(0, 60)}..."`);
            failed++;
          }
        }
      } catch (err) {
        failed++;
        errors.push(`Failed to apply edit: ${err instanceof Error ? err.message : String(err)}`);
      }
    }

    return { applied, failed, errors };
  });
}

export async function proposeDocumentEdits(
  host: OfficeHost,
  params: Record<string, unknown>,
): Promise<unknown> {
  if (host !== "word") {
    return { error: "office_propose_edits is only supported for Word documents." };
  }

  const edits = Array.isArray(params.edits) ? params.edits : [];
  const summary = typeof params.summary === "string" ? params.summary : "Proposed edits";

  if (!edits.length) {
    return { error: "No edits provided." };
  }

  return Word.run(async (context) => {
    const body = context.document.body;
    const paragraphs = body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();

    const docText = paragraphs.items.map((p) => p.text).join("\n");
    const verified: Array<Record<string, unknown>> = [];

    for (const edit of edits) {
      const record = edit as Record<string, unknown>;
      const kind = String(record.kind ?? "replace");
      const searchText = typeof record.searchText === "string" ? record.searchText : undefined;
      const oldText = typeof record.oldText === "string" ? record.oldText : searchText;
      const newText = typeof record.newText === "string" ? record.newText : "";
      const explanation = typeof record.explanation === "string" ? record.explanation : undefined;
      const id = typeof record.id === "string" ? record.id : crypto.randomUUID();

      let found = false;
      let contextPreview: string | undefined;
      if (oldText && docText.includes(oldText)) {
        found = true;
        const idx = docText.indexOf(oldText);
        const start = Math.max(0, idx - 40);
        const end = Math.min(docText.length, idx + oldText.length + 40);
        contextPreview = docText.slice(start, end);
      }

      verified.push({
        id,
        kind,
        searchText: searchText ?? oldText,
        oldText,
        newText: kind === "delete" ? "" : newText,
        explanation,
        found,
        contextPreview,
      });
    }

    return {
      proposalId: crypto.randomUUID(),
      summary,
      editCount: verified.length,
      edits: verified,
      instruction: "Present these proposed edits to the user for review. Each edit shows the text to find, the proposed change, and whether the target text was found in the document.",
    };
  });
}
