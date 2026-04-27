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
        const options = {
          ...(params.options && typeof params.options === "object" && !Array.isArray(params.options)
            ? params.options as Record<string, unknown>
            : {}),
          ...(params.font && typeof params.font === "object" && !Array.isArray(params.font)
            ? { font: params.font }
            : {}),
          style: params.style,
          builtInStyle: params.builtInStyle,
          alignment: params.alignment,
          leftIndent: params.leftIndent,
          rightIndent: params.rightIndent,
          firstLineIndent: params.firstLineIndent,
          lineSpacing: params.lineSpacing,
          spaceBefore: params.spaceBefore,
          spaceAfter: params.spaceAfter,
          clearFormatting: params.clearFormatting,
        };
        const result = await dependencies.applyHostAction(request.host, {
          type: "formatRange",
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

        const result = await dependencies.applyHostAction(request.host, {
          type: "formatList",
          target: request.params.target as never,
          options: {
            listType: request.params.listType ?? request.params.listKind,
            level: request.params.level,
            direction: request.params.direction,
            remove: request.params.remove,
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

        const result = await dependencies.applyHostAction(request.host, {
          type: "bookmark",
          target: request.params.target as never,
          options: {
            operation: request.params.operation,
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
          type: "hyperlink",
          target: request.params.target as never,
          content: typeof request.params.text === "string" ? request.params.text : undefined,
          options: {
            operation: request.params.operation,
            address: request.params.address,
            subAddress: request.params.subAddress,
            screenTip: request.params.screenTip,
            text: request.params.text,
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

        const result = await dependencies.applyHostAction(request.host, {
          type: "tableEdit",
          target: request.params.target as never,
          values: Array.isArray(request.params.values) ? request.params.values as never : undefined,
          options: {
            operation: request.params.operation,
            tableIndex: request.params.tableIndex,
            rowIndex: request.params.rowIndex,
            columnIndex: request.params.columnIndex,
            rowCount: request.params.rowCount,
            columnCount: request.params.columnCount,
            text: request.params.text,
            style: request.params.style,
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
          content: typeof request.params.content === "string" ? request.params.content : undefined,
          placement: typeof request.params.placement === "string" ? request.params.placement : undefined,
          options: {
            operation: request.params.operation,
            sectionIndex: request.params.sectionIndex,
            part: request.params.part,
            headerFooterType: request.params.headerFooterType,
            text: request.params.text,
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

        const result = await dependencies.applyHostAction(request.host, {
          type: "fieldReference",
          target: request.params.target as never,
          content: typeof request.params.text === "string" ? request.params.text : undefined,
          options: {
            operation: request.params.operation,
            fieldType: request.params.fieldType,
            fieldText: request.params.fieldText ?? request.params.text,
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

        const operation = String(request.params.operation ?? "");
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
            keepContent: request.params.keepContent,
            select: request.params.select,
            contentControlType: request.params.contentControlType,
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
            operation: request.params.operation,
            name: request.params.name,
            category: request.params.category,
            blockType: request.params.blockType,
            description: request.params.description,
            insertType: request.params.insertType,
            approved: request.params.approved,
            provenanceApproved: request.params.provenanceApproved,
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
          type: "annotationReview",
          target: request.params.target as never,
          content: typeof request.params.message === "string" ? request.params.message : undefined,
          options: {
            operation: request.params.operation,
            colorScheme: request.params.colorScheme,
            start: request.params.start,
            length: request.params.length,
            replacement: request.params.replacement,
            annotationId: request.params.annotationId,
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
