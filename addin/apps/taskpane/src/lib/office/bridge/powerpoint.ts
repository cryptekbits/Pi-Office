import type { OfficeToolRequest, OfficeToolResult } from "@pi-office/pi-office-pack/protocol";
import type { OfficeToolExecutorDependencies } from "./common";
import {
  firstString,
  toPayloadAwareResult,
  toPowerPointChartActionType,
  toPowerPointElementInsertActionType,
  toPowerPointElementRemoveActionType,
  toPowerPointMasterActionType,
  toPowerPointSlidesVerificationPayload,
  toPowerPointStructureActionType,
  toPowerPointStructureOptions,
  toPowerPointStructureTarget,
  toPowerPointTextAction,
  toPowerPointVisualVerificationPayload,
  toPowerPointXmlActionType,
  toStringArray,
  trimString,
} from "./common";

export async function executePowerPointOfficeTool(
  request: OfficeToolRequest,
  dependencies: OfficeToolExecutorDependencies,
): Promise<OfficeToolResult | undefined> {
      if (request.toolName === "get_presentation_structure") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_presentation_structure is only available for PowerPoint.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "getPresentationStructure",
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "get_slide") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "get_slide is only available for PowerPoint.",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "getSlide",
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "list_slide_shapes") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "list_slide_shapes is only available for PowerPoint.",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: "listSlideShapes",
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "modify_presentation_structure") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "modify_presentation_structure is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointStructureActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "modify_presentation_structure requires a supported operation (add_slide, move_slide, reorder_slides, delete_slide, apply_layout, select_slides, add_agenda_slide, add_transition_slide, combine_slides, import_slides_from_base64).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof request.params.content === "string" ? { content: request.params.content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "duplicate_slide") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "duplicate_slide is only available for PowerPoint.",
          };
        }

        const requestedSlideIds = toStringArray(request.params.slideIds);
        const fallbackSlideId = trimString(request.params.slideId);
        const slideIds = requestedSlideIds.length ? requestedSlideIds : fallbackSlideId ? [fallbackSlideId] : [];
        const targetParams =
          slideIds.length && !fallbackSlideId && typeof request.params.slideIndex !== "number"
            ? { ...request.params, slideId: slideIds[0] }
            : request.params;
        const target = toPowerPointStructureTarget(targetParams);
        const result = await dependencies.applyHostAction(request.host, {
          type: "duplicateSlide",
          ...(target ? { target } : {}),
          options: {
            ...toPowerPointStructureOptions(request.params),
            ...(slideIds.length ? { slideIds } : {}),
          },
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "insert_slide_element") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "insert_slide_element is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointElementInsertActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "insert_slide_element requires a supported operation (add_text_box, add_geometric_shape, add_table, add_line, add_process_flow, add_simple_diagram, insert_inline_picture).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.text, request.params.base64);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "remove_slide_element") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "remove_slide_element is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointElementRemoveActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "remove_slide_element requires a supported operation (remove_shape, remove_shapes, clear_shape_text).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_text") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_text is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const textAction = toPowerPointTextAction(operation, request.params);
        if (!textAction) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_text requires a supported operation (set_shape_text, append_shape_text, clear_shape_text, insert_text).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.text, request.params.newText) ?? "";
        const actionOptions = toPowerPointStructureOptions(request.params);
        const placement = trimString(request.params.placement) ?? textAction.placement;
        if (placement && !("placement" in actionOptions)) {
          actionOptions.placement = placement;
        }
        const result = await dependencies.applyHostAction(request.host, {
          type: textAction.type,
          ...(target ? { target } : {}),
          ...(textAction.type !== "clearShapeText" ? { content } : {}),
          ...(placement ? { placement } : {}),
          options: actionOptions,
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_xml") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_xml is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointXmlActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_xml requires a supported operation (inspect_presentation_package, get_presentation_theme, get_slide_notes, set_slide_notes, replace_slide_notes, import_slides_from_base64, merge_presentation_from_base64, export_slides_as_base64).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.text, request.params.base64);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_master") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_master is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointMasterActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_master is a legacy-named layout application tool. It currently supports only apply_layout/set_layout; it does not edit slide masters.",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "edit_slide_chart") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "edit_slide_chart is only available for PowerPoint.",
          };
        }

        const operation =
          trimString(request.params.operation) ??
          trimString(request.params.mode) ??
          trimString(request.params.type);
        const actionType = toPowerPointChartActionType(operation);
        if (!actionType) {
          return {
            requestId: request.requestId,
            success: false,
            error:
              "edit_slide_chart requires a supported operation (get_slide_charts, add_slide_chart, update_slide_chart).",
          };
        }

        const target = toPowerPointStructureTarget(request.params);
        const content = firstString(request.params.content, request.params.title);
        const result = await dependencies.applyHostAction(request.host, {
          type: actionType,
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "copy_image_between_slides") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "copy_image_between_slides is only available for PowerPoint.",
          };
        }

        const targetParams: Record<string, unknown> = {
          ...request.params,
          slideId: trimString(request.params.targetSlideId) ?? trimString(request.params.slideId),
          slideIndex:
            typeof request.params.targetSlideIndex === "number"
              ? request.params.targetSlideIndex
              : request.params.slideIndex,
          shapeId: trimString(request.params.targetShapeId) ?? trimString(request.params.shapeId),
        };
        const target = toPowerPointStructureTarget(targetParams);
        const content = firstString(request.params.sourceImageBase64, request.params.base64, request.params.content);
        const result = await dependencies.applyHostAction(request.host, {
          type: "copyImageBetweenSlides",
          ...(target ? { target } : {}),
          ...(typeof content === "string" ? { content } : {}),
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "search_icons") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_icons is only available for PowerPoint.",
          };
        }

        const query = firstString(request.params.query, request.params.search, request.params.content);
        if (!query?.trim()) {
          return {
            requestId: request.requestId,
            success: false,
            error: "search_icons requires a non-empty query.",
          };
        }

        const result = await dependencies.applyHostAction(request.host, {
          type: "searchIcons",
          content: query,
          options: {
            ...toPowerPointStructureOptions(request.params),
            query,
          },
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "insert_icon") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "insert_icon is only available for PowerPoint.",
          };
        }

        const iconId = firstString(request.params.iconId, request.params.iconName, request.params.query, request.params.content);
        if (!iconId?.trim()) {
          return {
            requestId: request.requestId,
            success: false,
            error: "insert_icon requires iconId, iconName, query, or content.",
          };
        }

        const targetParams: Record<string, unknown> = {
          ...request.params,
          slideId: trimString(request.params.targetSlideId) ?? trimString(request.params.slideId),
          slideIndex:
            typeof request.params.targetSlideIndex === "number"
              ? request.params.targetSlideIndex
              : request.params.slideIndex,
          shapeId: trimString(request.params.targetShapeId) ?? trimString(request.params.shapeId),
        };
        const target = toPowerPointStructureTarget(targetParams);
        const result = await dependencies.applyHostAction(request.host, {
          type: "insertIcon",
          ...(target ? { target } : {}),
          content: iconId,
          options: toPowerPointStructureOptions(request.params),
        });
        return toPayloadAwareResult(request.requestId, result);
      }

      if (request.toolName === "verify_slides") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_slides is only available for PowerPoint.",
          };
        }

        const scope = trimString(request.params.scope);
        const result = await dependencies.applyHostAction(request.host, {
          type: "getPresentationStructure",
          options: toPowerPointStructureOptions(request.params),
        });
        const payloadAware = toPayloadAwareResult(request.requestId, result);
        if (!payloadAware.success) {
          return payloadAware;
        }

        return {
          requestId: request.requestId,
          success: true,
          content: toPowerPointSlidesVerificationPayload(payloadAware.content, scope),
        };
      }

      if (request.toolName === "verify_slide_visual") {
        if (request.host !== "powerpoint") {
          return {
            requestId: request.requestId,
            success: false,
            error: "verify_slide_visual is only available for PowerPoint.",
          };
        }

        const includeFormatting = request.params.includeFormatting !== false;
        const maxImages =
          typeof request.params.maxImages === "number" && Number.isFinite(request.params.maxImages)
            ? Math.max(1, Math.min(4, Math.trunc(request.params.maxImages)))
            : 1;
        const scope = trimString(request.params.scope) ?? "slide";
        const rawResult = await dependencies.collectOfficeContext(request.host, {
          includeFormatting,
          maxImages,
          scope,
        });
        const payloadAware = toPayloadAwareResult(request.requestId, rawResult);
        if (!payloadAware.success) {
          return payloadAware;
        }

        return {
          requestId: request.requestId,
          success: true,
          content: toPowerPointVisualVerificationPayload(payloadAware.content, maxImages),
        };
      }


  return undefined;
}
