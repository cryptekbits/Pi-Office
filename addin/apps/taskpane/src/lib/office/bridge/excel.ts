import type { OfficeToolRequest, OfficeToolResult } from "@pi-office/pi-office-pack/protocol";
import type { OfficeToolExecutorDependencies } from "./common";
import {
  excelSelectionMatchesRequest,
  firstString,
  formatExcelRangeHint,
  isRecord,
  isSingleCellAddress,
  serializeExcelInventoryEntry,
  toAnchor,
  toCsv,
  toExcelInventoryEntries,
  toExcelMatrixFromPayload,
  toExcelObjectActionType,
  toExcelObjectKindSet,
  toExcelObjectTarget,
  toExcelRangeTarget,
  toExcelSheetStructureActionType,
  toExcelToolOptions,
  toHostAction,
  toPayloadAwareResult,
  trimString,
} from "./common";

export async function executeExcelOfficeTool(
  request: OfficeToolRequest,
  dependencies: OfficeToolExecutorDependencies,
): Promise<OfficeToolResult | undefined> {
      if (request.toolName === "get_cell_ranges") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_cell_ranges is only available for Excel.",
          };
        }

        const target = toExcelRangeTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "getRangeValues",
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["anchor", "sheetName", "address"]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "set_cell_range") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "set_cell_range is only available for Excel.",
          };
        }

        const action = toHostAction({
          ...request.params,
          operation: "setRangeValues",
          mode: "setRangeValues",
          format: "matrix",
        });
        const result = await dependencies.applyHostAction(request.host, {
          ...action,
          type: "setRangeValues",
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "clear_cell_range") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "clear_cell_range is only available for Excel.",
          };
        }

        const target = toExcelRangeTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "clearRange",
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["anchor", "sheetName", "address"]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "resize_range") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "resize_range is only available for Excel.",
          };
        }

        const target = toExcelRangeTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "resizeRange",
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["anchor", "sheetName", "address"]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "copy_to") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "copy_to is only available for Excel.",
          };
        }

        const sourceTarget =
          toExcelRangeTarget(request.params, {
            sheetNameKey: "sourceSheetName",
            addressKey: "sourceAddress",
            anchorKey: "sourceAnchor",
          }) ?? toExcelRangeTarget(request.params);
        const destinationTarget = toExcelRangeTarget(request.params, {
          sheetNameKey: "destinationSheetName",
          addressKey: "destinationAddress",
          anchorKey: "destinationAnchor",
        });
        if (!destinationTarget?.address) {
          return {
            requestId: request.requestId,
            success: false,
            error: "copy_to requires destinationAddress (and optional destinationSheetName).",
          };
        }

        const options = toExcelToolOptions(request.params, [
          "anchor",
          "sheetName",
          "address",
          "sourceAnchor",
          "sourceSheetName",
          "sourceAddress",
          "destinationAnchor",
          "destinationSheetName",
          "destinationAddress",
        ]);
        options.destinationAddress = destinationTarget.address;
        if (destinationTarget.sheetName) {
          options.destinationSheetName = destinationTarget.sheetName;
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "copyRange",
          ...(sourceTarget ? { target: sourceTarget } : {}),
          options,
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "modify_sheet_structure") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "modify_sheet_structure is only available for Excel.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toExcelSheetStructureActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "modify_sheet_structure requires a supported operation (create_worksheet, rename_worksheet, duplicate_worksheet, delete_worksheet).",
          };
        }

        const sheetName = trimString(request.params.sheetName) ?? trimString(request.params.sourceSheetName);
        const anchor = isRecord(request.params.anchor) && Object.keys(request.params.anchor).length > 0
          ? request.params.anchor
          : undefined;
        const target = sheetName || anchor ? toAnchor({
          ...(anchor ? { anchor } : {}),
          ...(sheetName ? { sheetName } : {}),
        }) : undefined;

        if (actionType !== "createWorksheet" && !target) {
          return {
            requestId: request.requestId,
            success: false,
            error: "modify_sheet_structure requires sheetName or anchor for rename, duplicate, and delete operations.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["operation", "mode", "type", "sheetName", "sourceSheetName", "anchor"]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "modify_object") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "modify_object is only available for Excel.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toExcelObjectActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "modify_object requires a supported operation (format_range, create_table, format_table, apply_table_filter, clear_table_filter, clear_table_filters, reapply_table_filters, create_chart, update_chart, create_pivot_table, update_pivot_table, sort_pivot_field, sort_pivot_by_labels, sort_pivot_by_values, refresh_pivot_table, set_worksheet_gridlines, set_worksheet_headings, set_print_area, set_data_validation, clear_data_validation, add_conditional_format, clear_conditional_formats, insert_inline_picture).",
          };
        }

        const target = toExcelObjectTarget(request.params);
        const content = firstString(
          request.params.content,
          request.params.text,
          request.params.base64,
          request.params.imageBase64,
        );
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toExcelToolOptions(request.params, [
            "operation",
            "mode",
            "type",
            "anchor",
            "sheetName",
            "address",
            "tableName",
            "chartName",
            "pivotTableName",
            "namedItemName",
            "content",
            "text",
            "base64",
            "imageBase64",
          ]),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "get_all_objects") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_all_objects is only available for Excel.",
          };
        }

        const scope = trimString(request.params.scope) ?? "workbook";
        const includeFormatting = request.params.includeFormatting !== false;
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages: 0,
          scope,
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const requestedKinds = toExcelObjectKindSet(request.params.objectTypes);
        const entries = toExcelInventoryEntries(payloadAware.content).filter((entry) =>
          requestedKinds ? requestedKinds.has(entry.kind) : true,
        );
        const limit =
          typeof request.params.limit === "number" && Number.isFinite(request.params.limit)
            ? Math.max(1, Math.min(200, Math.trunc(request.params.limit)))
            : entries.length || 200;
        const limited = entries.slice(0, limit);

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary: `Excel object inventory captured (${limited.length} item${limited.length === 1 ? "" : "s"}).`,
            details: {
              kind: "excel-object-inventory",
              mutating: false,
              scope,
              requestedObjectTypes: requestedKinds ? Array.from(requestedKinds) : undefined,
              matchCount: limited.length,
              totalCount: entries.length,
              objects: {
                tables: limited.filter((entry) => entry.kind === "table").map((entry) => serializeExcelInventoryEntry(entry)),
                charts: limited.filter((entry) => entry.kind === "chart").map((entry) => serializeExcelInventoryEntry(entry)),
                pivotTables: limited.filter((entry) => entry.kind === "pivotTable").map((entry) => serializeExcelInventoryEntry(entry)),
                namedItems: limited.filter((entry) => entry.kind === "namedItem").map((entry) => serializeExcelInventoryEntry(entry)),
                worksheets: limited.filter((entry) => entry.kind === "worksheet").map((entry) => serializeExcelInventoryEntry(entry)),
                cells: limited.filter((entry) => entry.kind === "cell").map((entry) => serializeExcelInventoryEntry(entry)),
              },
            },
          },
        };
      }

      if (request.toolName === "search_data") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_data is only available for Excel.",
          };
        }

        const query = trimString(request.params.query) ?? trimString(request.params.search) ?? trimString(request.params.text);
        if (!query) {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_data requires a non-empty query.",
          };
        }

        const scope = trimString(request.params.scope) ?? "workbook";
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting: false,
          maxImages: 0,
          scope,
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const requestedKinds = toExcelObjectKindSet(request.params.objectTypes);
        const queryLower = query.toLowerCase();
        const matches = toExcelInventoryEntries(payloadAware.content)
          .filter((entry) => (requestedKinds ? requestedKinds.has(entry.kind) : true))
          .filter((entry) => {
            const fields = [entry.label, entry.name, entry.sheetName, entry.address, entry.type, entry.text, entry.formula]
              .filter((value): value is string => typeof value === "string");
            return fields.some((value) => value.toLowerCase().includes(queryLower));
          });

        const limit =
          typeof request.params.limit === "number" && Number.isFinite(request.params.limit)
            ? Math.max(1, Math.min(200, Math.trunc(request.params.limit)))
            : 40;
        const limited = matches.slice(0, limit);

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary: `Excel search found ${limited.length} match${limited.length === 1 ? "" : "es"} for "${query}".`,
            details: {
              kind: "excel-data-search",
              mutating: false,
              query,
              scope,
              requestedObjectTypes: requestedKinds ? Array.from(requestedKinds) : undefined,
              matchCount: limited.length,
              totalMatches: matches.length,
              matches: limited.map((entry) => serializeExcelInventoryEntry(entry)),
            },
          },
        };
      }

      if (request.toolName === "get_range_as_csv") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_range_as_csv is only available for Excel.",
          };
        }

        const target = toExcelRangeTarget(request.params);
        const includeFormulas = request.params.includeFormulas === true;
        const delimiter = trimString(request.params.delimiter) ?? ",";
        const quoteValues = request.params.quoteValues === true;
        const includeHeaders = request.params.includeHeaders !== false;
        const rawResult = await dependencies.applyHostAction(request.host, {
          type: "getRangeValues",
          ...(target ? { target } : {}),
          options: {
            ...toExcelToolOptions(request.params, [
              "anchor",
              "sheetName",
              "address",
              "delimiter",
              "quoteValues",
              "includeHeaders",
              "includeFormulas",
            ]),
            includeValues: true,
            includeText: true,
            includeFormulas: true,
          },
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const matrix = toExcelMatrixFromPayload(payloadAware.content, includeFormulas);
        const rows = includeHeaders ? matrix : matrix.slice(1);
        const csv = toCsv(rows, delimiter, quoteValues);
        const payloadRecord = isRecord(payloadAware.content) ? payloadAware.content : {};
        const payloadData = isRecord(payloadRecord.data) ? payloadRecord.data : payloadRecord;
        const sheetName = trimString(payloadData.sheetName) ?? target?.sheetName;
        const address = trimString(payloadData.address) ?? target?.address;
        const columnCount = rows.length && Array.isArray(rows[0]) ? rows[0].length : 0;

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary: `Exported ${rows.length} row${rows.length === 1 ? "" : "s"} from Excel range as CSV.`,
            csv,
            details: {
              kind: "excel-range-csv-export",
              mutating: false,
              sheetName,
              address,
              rowCount: rows.length,
              columnCount,
              delimiter,
              includeFormulas,
            },
          },
        };
      }

      if (request.toolName === "read_range_image") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "read_range_image is only available for Excel.",
          };
        }

        const includeFormatting = request.params.includeFormatting !== false;
        const maxImages =
          typeof request.params.maxImages === "number" && Number.isFinite(request.params.maxImages)
            ? Math.max(1, Math.min(4, Math.trunc(request.params.maxImages)))
            : 1;
        const scope = trimString(request.params.scope) ?? "selection";
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages,
          scope,
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const payloadRecord = isRecord(payloadAware.content) ? payloadAware.content : {};
        const visuals = Array.isArray(payloadRecord.visuals) ? payloadRecord.visuals : [];
        const requestedSheetName = trimString(request.params.sheetName);
        const requestedAddress = trimString(request.params.address);
        const state = isRecord(payloadRecord.state) ? payloadRecord.state : {};
        const selection = isRecord(state.selection) ? state.selection : {};
        const activeSelectionLabel = trimString(selection.label);
        const requestedRangeHonored = excelSelectionMatchesRequest(
          activeSelectionLabel,
          requestedSheetName,
          requestedAddress,
        );
        if (requestedRangeHonored === false) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              `read_range_image can only capture the active Excel selection as an Office.js image snapshot. ` +
              `The active selection is ${activeSelectionLabel ?? "unknown"}, not ${formatExcelRangeHint(requestedSheetName, requestedAddress)}. ` +
              "Navigate to or select the target range first, then retry.",
          };
        }
        if (!visuals.length) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "read_range_image could not capture an Office.js image snapshot of the active Excel selection. " +
              "Browser-only runtime does not render arbitrary ranges by address; select the target range first and retry.",
          };
        }

        const captureNote =
          requestedRangeHonored === true
            ? "Captured the active Excel selection through Office.js image coercion."
            : "Captured the active Excel selection through Office.js image coercion; requested sheet/address values are advisory unless they match the active selection.";

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary:
              trimString(payloadRecord.summary) ??
              `Excel active-selection visual snapshot captured (${visuals.length} image${visuals.length === 1 ? "" : "s"}). ${captureNote}`,
            visual: {
              kind: "excel-selection-snapshot",
              captureMode: "officejs-selection-snapshot",
              imageCount: visuals.length,
              scopeRequested: scope,
              requestedRangeHonored,
              note: captureNote,
            },
            details: {
              kind: "excel-selection-snapshot-read",
              mutating: false,
              host: "excel",
              scope,
              requestedRange: {
                sheetName: requestedSheetName,
                address: requestedAddress,
                honored: requestedRangeHonored,
              },
              selection,
              formatting: payloadRecord.formatting,
            },
            visuals,
          },
        };
      }

      if (request.toolName === "extract_chart_xml") {
        if (request.host !== "excel") {
          return {
            requestId: request.requestId,
            success: false,
            error: "extract_chart_xml is only available for Excel.",
          };
        }

        const chartName = trimString(request.params.chartName) ?? trimString(request.params.name);
        const chartId = trimString(request.params.chartId) ?? trimString(request.params.id);
        const chartIndex =
          typeof request.params.chartIndex === "number" && Number.isFinite(request.params.chartIndex)
            ? Math.max(1, Math.trunc(request.params.chartIndex))
            : undefined;
        const hasAnchor = isRecord(request.params.anchor) && Object.keys(request.params.anchor).length > 0;
        if (!chartName && !chartId && typeof chartIndex !== "number" && !hasAnchor) {
          return {
            requestId: request.requestId,
            success: false,
            error: "extract_chart_xml requires chartName, chartId, chartIndex, or a chart anchor.",
          };
        }

        const target = toExcelObjectTarget({
          ...request.params,
          ...(chartName ? { chartName } : {}),
          ...(chartId ? { id: chartId } : {}),
        });
        const rawResult = await dependencies.applyHostAction(request.host, {
          type: "extractChartXml",
          ...(target ? { target } : {}),
          options: toExcelToolOptions(request.params, ["anchor", "sheetName", "chartName", "chartId", "chartIndex"]),
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const payloadRecord = isRecord(payloadAware.content) ? payloadAware.content : {};
        const payloadData = isRecord(payloadRecord.data) ? payloadRecord.data : payloadRecord;
        const xml = trimString(payloadData.chartXml) ?? trimString(payloadData.xml) ?? trimString(payloadRecord.chartXml);
        if (!xml) {
          return {
            requestId: request.requestId,
            success: false,
            error: "extract_chart_xml did not return chart XML content.",
          };
        }

        const resolvedChartName = trimString(payloadData.chartName) ?? chartName;
        const resolvedSheetName = trimString(payloadData.sheetName) ?? trimString(request.params.sheetName) ?? target?.sheetName;

        return {
          requestId: request.requestId,
          success: true,
          content: {
            summary: trimString(payloadRecord.summary) ?? `Extracted chart XML for ${resolvedChartName ?? "the target chart"}.`,
            xml,
            details: {
              kind: "excel-chart-xml",
              mutating: false,
              chartName: resolvedChartName,
              sheetName: resolvedSheetName,
              extraction: "runtime-generated-chart-metadata-xml",
            },
          },
        };
      }


  return undefined;
}
