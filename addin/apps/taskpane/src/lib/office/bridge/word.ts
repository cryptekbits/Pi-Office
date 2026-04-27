import type { OfficeToolRequest, OfficeToolResult } from "@pi-office/pi-office-pack/protocol";
import type { OfficeToolExecutorDependencies } from "./common";
import {
  appendViewportCaptureSummary,
  toHostAction,
  toPayloadAwareResult,
  toWordDocumentVerificationPayload,
  toWordVisualVerificationPayload,
  trimString,
} from "./common";

function toNumber(value: unknown, fallback: number): number {
  return typeof value === "number" && Number.isFinite(value) ? value : fallback;
}

function toOptionalNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

function toZeroBasedIndex(value: unknown): number | undefined {
  const number = toOptionalNumber(value);
  return typeof number === "number" ? Math.max(0, Math.trunc(number) - 1) : undefined;
}

function toStringArray(value: unknown): string[] {
  return Array.isArray(value) ? value.map((entry) => String(entry)).filter(Boolean) : [];
}

export async function executeWordOfficeTool(
  request: OfficeToolRequest,
  dependencies: OfficeToolExecutorDependencies,
): Promise<OfficeToolResult | undefined> {
      if (request.toolName === "edit_doc_text") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_doc_text is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, toHostAction(request.params));
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "edit_doc_list") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_doc_list is only available for Word.",
          };
        }

        const result = await dependencies.proposeEdits(request.host, request.params);
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "word_format_text") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_format_text is only available for Word.",
          };
        }

        const params = request.params as Record<string, unknown>;
        const font = params.font && typeof params.font === "object" && !Array.isArray(params.font)
          ? params.font as Record<string, unknown>
          : {};
        const paragraph = params.paragraph && typeof params.paragraph === "object" && !Array.isArray(params.paragraph)
          ? params.paragraph as Record<string, unknown>
          : {};
        const options = {
          ...(params.options && typeof params.options === "object" && !Array.isArray(params.options)
            ? params.options as Record<string, unknown>
            : {}),
          style: params.style,
          styleBuiltIn: params.styleBuiltIn ?? params.builtInStyle,
          alignment: params.alignment,
          fontName: font.name ?? params.fontName,
          fontSize: font.size ?? params.fontSize,
          color: font.color ?? params.color,
          bold: font.bold ?? params.bold,
          italic: font.italic ?? params.italic,
          highlightColor: font.highlightColor ?? params.highlightColor,
          leftIndent: paragraph.leftIndent ?? params.leftIndent,
          rightIndent: paragraph.rightIndent ?? params.rightIndent,
          firstLineIndent: paragraph.firstLineIndent ?? params.firstLineIndent,
          lineSpacing: paragraph.lineSpacing ?? params.lineSpacing,
          spaceBefore: paragraph.spaceBefore ?? params.spaceBefore,
          spaceAfter: paragraph.spaceAfter ?? params.spaceAfter,
          clearFormatting: params.clearFormatting,
        };
        const result = await dependencies.applyHostAction(request.host, {
          type: request.params.clearFormatting === true ? "clearFormatting" : "applyTextFormat",
          target: params.target as never,
          options,
        });
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "word_search") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_search is only available for Word.",
          };
        }

        if (!dependencies.searchWordDocument) {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_search is not available in this taskpane build.",
          };
        }

        const query = trimString(request.params.query);
        if (!query) {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_search requires query.",
          };
        }

        const result = await dependencies.searchWordDocument({
          query,
          objectTypes: toStringArray(request.params.objectTypes),
          maxResults: toNumber(request.params.maxResults ?? request.params.limit, 20),
          includeContext: request.params.includeContext !== false,
          matchCase: request.params.matchCase === true,
          matchWholeWord: request.params.matchWholeWord === true,
        });
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "word_list_format") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_list_format is only available for Word.",
          };
        }

        const operation = String(request.params.operation ?? request.params.listType ?? request.params.listKind ?? "bullet");
        const listKindMap: Record<string, string> = {
          bullet: "bullet",
          bullets: "bullet",
          number: "numbered",
          numbered: "numbered",
          outlineNumber: "outline",
          outline: "outline",
          multilevel: "outline",
          removeNumbers: "remove",
          remove: "remove",
          none: "remove",
          setLevel: "numbered",
        };
        if (operation === "indent" || operation === "outdent") {
          return {
            requestId: request.requestId,
            success: false,
            error: `word_list_format operation ${operation} is not implemented by the current Word list executor; use setLevel with an explicit level instead.`,
          };
        }
        const result = await dependencies.applyHostAction(request.host, {
          type: "applyListFormat",
          target: request.params.target as never,
          options: {
            listKind: listKindMap[operation] ?? request.params.listType ?? request.params.listKind ?? operation,
            level: request.params.level,
            confirmed: request.params.confirmed ?? request.params.confirmBroadChange,
          },
        });
        return {
          requestId: request.requestId,
          success: true,
          content: result,
        };
      }

      if (request.toolName === "word_reference_inventory") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_reference_inventory is only available for Word.",
          };
        }

        const operation = String(request.params.operation ?? "inventory");
        if (operation === "add" || operation === "delete") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_reference_inventory is read-only. Bookmark add/delete needs a separate write-doc tool before it can be exposed.",
          };
        }
        const result = await dependencies.applyHostAction(request.host, {
          type: "bookmarkAction",
          target: request.params.target as never,
          options: {
            operation,
            name: request.params.name,
            preserve: request.params.preserve,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_hyperlink") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_hyperlink is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "manageHyperlink",
          target: request.params.target as never,
          content: typeof request.params.textToDisplay === "string"
            ? request.params.textToDisplay
            : typeof request.params.text === "string"
              ? request.params.text
              : undefined,
          options: {
            operation: request.params.operation,
            address: request.params.address,
            subAddress: request.params.subAddress,
            screenTip: request.params.screenTip,
            textToDisplay: request.params.textToDisplay ?? request.params.text,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_table") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_table is only available for Word.",
          };
        }

        const operation = String(request.params.operation ?? "inventory");
        const type = operation === "inventory"
          ? "inspectTable"
          : operation === "setCellText"
            ? "editTableCell"
            : "modifyTable";
        const tableOperationMap: Record<string, string> = {
          formatTable: "setStyle",
          deleteTable: "deleteTable",
          addRows: "addRows",
          addColumns: "addColumns",
          deleteRows: "deleteRows",
          deleteColumns: "deleteColumns",
          clear: "clear",
        };
        if (operation === "mergeCells") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_table mergeCells is not implemented by the current Word table executor.",
          };
        }
        const result = await dependencies.applyHostAction(request.host, {
          type,
          target: request.params.target as never,
          values: Array.isArray(request.params.values) ? request.params.values as never : undefined,
          options: {
            operation: tableOperationMap[operation] ?? operation,
            tableIndex: toZeroBasedIndex(request.params.tableIndex),
            rowIndex: toZeroBasedIndex(request.params.rowIndex),
            columnIndex: toZeroBasedIndex(request.params.columnIndex),
            rowCount: request.params.rowCount,
            columnCount: request.params.columnCount,
            count: request.params.count,
            text: request.params.text,
            style: request.params.style,
            styleBuiltIn: request.params.styleBuiltIn,
            shadingColor: request.params.shadingColor,
            alignment: request.params.alignment,
            confirmDestructive: request.params.confirmDestructive,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_section_layout") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_section_layout is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "sectionLayout",
          target: request.params.target as never,
          content: typeof request.params.text === "string"
            ? request.params.text
            : typeof request.params.content === "string"
              ? request.params.content
              : undefined,
          placement: typeof request.params.placement === "string" ? request.params.placement : undefined,
          options: {
            operation: request.params.operation,
            sectionIndex: toZeroBasedIndex(request.params.sectionIndex),
            part: request.params.part,
            headerFooterType: request.params.headerFooterType,
            text: request.params.text,
            confirmDestructive: request.params.confirmDestructive,
            margins: request.params.margins,
            topMargin: request.params.topMargin,
            bottomMargin: request.params.bottomMargin,
            leftMargin: request.params.leftMargin,
            rightMargin: request.params.rightMargin,
            pageWidth: request.params.pageWidth,
            pageHeight: request.params.pageHeight,
            orientation: request.params.orientation,
            breakType: request.params.breakType,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_field_reference") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_field_reference is only available for Word.",
          };
        }

        const operation = String(request.params.operation ?? "inventory");
        const fieldTarget = request.params.target ?? (typeof request.params.fieldId === "string"
          ? { kind: "field", id: request.params.fieldId }
          : undefined);
        const fieldOperationMap: Record<string, { type: string; operation?: string }> = {
          inventory: { type: "fieldAction", operation: "inventory" },
          insertField: { type: "insertField" },
          updateField: { type: "fieldAction", operation: "update" },
          lockField: { type: "fieldAction", operation: "lock" },
          unlockField: { type: "fieldAction", operation: "unlock" },
          unlinkField: { type: "fieldAction", operation: "unlink" },
          selectField: { type: "fieldAction", operation: "select" },
          tocInventory: { type: "tocAction", operation: "inventory" },
          updateTocPageNumbers: { type: "tocAction", operation: "updatePageNumbers" },
        };
        const mapped = fieldOperationMap[operation] ?? { type: "fieldAction", operation };
        const result = await dependencies.applyHostAction(request.host, {
          type: mapped.type,
          target: fieldTarget as never,
          content: typeof request.params.text === "string" ? request.params.text : undefined,
          options: {
            operation: mapped.operation,
            fieldType: request.params.fieldType,
            fieldText: request.params.fieldText ?? request.params.text,
            text: request.params.fieldText ?? request.params.text,
            lock: request.params.lock,
            confirmDestructive: request.params.confirmDestructive,
            tocIndex: request.params.tocIndex,
            lowerHeadingLevel: request.params.lowerHeadingLevel,
            upperHeadingLevel: request.params.upperHeadingLevel,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_content_control") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_content_control is only available for Word.",
          };
        }

        const requestedOperation = String(request.params.operation ?? "");
        const operation =
          requestedOperation === "deleteWrapper" || requestedOperation === "deleteContent"
            ? "delete"
            : requestedOperation === "lock" || requestedOperation === "unlock"
              ? "setMetadata"
              : requestedOperation;
        if (operation === "inventory") {
          const rawResult = await dependencies.collectOfficeContext(request.host, {
            includeFormatting: true,
            maxImages: 0,
            scope: "contentControls",
          });
          const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
          if (!payloadAware.success) return payloadAware;
          return {
            requestId: request.requestId,
            success: true,
            content: (payloadAware.content as { snippets?: { contentControls?: unknown } })?.snippets?.contentControls ?? [],
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "contentControlEdit",
          target: request.params.target as never,
          content: typeof request.params.text === "string" ? request.params.text : undefined,
          options: {
            operation,
            title: request.params.title,
            tag: request.params.tag,
            placeholderText: request.params.placeholderText,
            appearance: request.params.appearance,
            cannotDelete: request.params.cannotDelete,
            cannotEdit: request.params.cannotEdit,
            removeWhenEdited: request.params.removeWhenEdited,
            keepContent: requestedOperation === "deleteWrapper" ? true : request.params.keepContent,
            confirmDestructive: request.params.confirmDestructive,
            select: request.params.select,
            contentControlType: request.params.contentControlType,
            ...(requestedOperation === "lock" ? { cannotEdit: true, cannotDelete: true } : {}),
            ...(requestedOperation === "unlock" ? { cannotEdit: false, cannotDelete: false } : {}),
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_building_block") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_building_block is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "buildingBlock",
          target: request.params.target as never,
          content: typeof request.params.content === "string" ? request.params.content : undefined,
          placement: typeof request.params.placement === "string" ? request.params.placement : undefined,
          options: {
            operation: request.params.operation === "insertApprovedText" && typeof request.params.text === "string"
              ? "insertApprovedText"
              : request.params.operation === "insertApprovedText"
                ? "insert"
                : request.params.operation,
            name: request.params.name,
            category: request.params.category,
            blockType: request.params.blockType,
            description: request.params.description,
            insertType: request.params.insertType,
            approved: request.params.approved,
            provenanceApproved: request.params.provenanceApproved,
            provenance: request.params.provenance,
            text: request.params.text,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_annotation_review") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_annotation_review is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "critiqueAnnotation",
          target: request.params.target as never,
          content: typeof request.params.message === "string" ? request.params.message : undefined,
          options: {
            operation: request.params.operation === "insert" || request.params.operation === "fallbackProposal"
              ? "propose"
              : request.params.operation,
            colorScheme: request.params.colorScheme,
            start: request.params.start,
            length: request.params.length,
            replacement: request.params.replacement,
            annotationId: request.params.annotationId,
            edits: request.params.fallbackEdits,
            critiques: request.params.critiques,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_redline_review") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_redline_review is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "reviewExchange",
          target: request.params.target as never,
          content: typeof request.params.baselinePath === "string" ? request.params.baselinePath : undefined,
          options: {
            operation: request.params.operation === "diagnostics" ? "inventory" : request.params.operation,
            baselinePath: request.params.baselinePath,
            filePath: request.params.baselinePath,
            compareTarget: request.params.compareTarget,
            confirmReviewStateChange: request.params.confirmReviewStateChange,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_proofing_stats") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_proofing_stats is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "proofingStats",
          target: request.params.target as never,
          options: {
            scope: request.params.scope,
            includeReadability: request.params.includeReadability,
            includeProofing: request.params.includeProofing,
            includeCounts: request.params.includeCounts,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }

      if (request.toolName === "word_collab_guard") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "word_collab_guard is only available for Word.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "protectionAwareness",
          target: request.params.target as never,
          options: {
            reviewerName: request.params.reviewerName,
            visible: request.params.visible,
            includeReviewers: request.params.includeReviewers,
            includeRevisions: request.params.includeRevisions,
          },
        });
        return { requestId: request.requestId, success: true, content: result };
      }


      if (request.toolName === "verify_doc") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_doc is only available for Word.",
          };
        }

        const scope = trimString(request.params.scope);
        const includeFormatting = request.params.includeFormatting !== false;
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages: 0,
          ...(scope ? { scope } : {}),
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        return {
          requestId: request.requestId,
          success: true,
          content: toWordDocumentVerificationPayload(payloadAware.content, scope),
        };
      }

      if (request.toolName === "verify_doc_visual") {
        if (request.host !== "word") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_doc_visual is only available for Word.",
          };
        }

        const includeFormatting = request.params.includeFormatting !== false;
        const includeWindowFrameRequested = request.params.includeWindowFrame === true;
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages: 1,
          scope: "viewport",
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        const viewportPayload = appendViewportCaptureSummary(payloadAware.content, includeWindowFrameRequested);
        return {
          requestId: request.requestId,
          success: true,
          content: toWordVisualVerificationPayload(viewportPayload, includeWindowFrameRequested),
        };
      }


      if (request.toolName === "office_propose_edits") {
        const result = await dependencies.proposeEdits(request.host, request.params);
        return toPayloadAwareResult(request.requestId, result);
      }


  return undefined;
}
