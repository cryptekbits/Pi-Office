import type { OfficeHostAction } from "@pi-office/pi-office-pack/protocol";
import { resolveExcelWorksheet } from "../excel-targets";

export function isExcelMediaAction(type: string): boolean {
  return type === "insertInlinePicture";
}

export async function applyExcelMediaAction(
  context: Excel.RequestContext,
  action: OfficeHostAction,
  type: string,
): Promise<unknown> {
  if (type === "insertInlinePicture") {
    const worksheet = resolveExcelWorksheet(context, action.target, true);
    const imageBase64 = action.content ?? "";
    const shape = worksheet.shapes.addImage(`data:image/png;base64,${imageBase64}`);
    shape.load("id,name,width,height");
    await context.sync();
    return { ok: true, host: "excel", action: type, shapeId: shape.id, shapeName: shape.name, width: shape.width, height: shape.height };
  }

  throw new Error(`Unsupported Excel media action: ${type}`);
}
