import type { WorkbookDefinition } from "../core/types";
import { validateWorkbookDefinition } from "../core/validation";
import { createSheetRegistry } from "../core/sheet-registry";
import type { RenderPlan } from "./render-plan";
import { RenderPlanBuilder } from "./render-plan-builder";

export function compileWorkbookToRenderPlan(workbook: WorkbookDefinition): RenderPlan {
  validateWorkbookDefinition(workbook);
  createSheetRegistry(workbook.sheets);

  const builder = new RenderPlanBuilder(workbook.metadata, workbook.styles);

  for (const sheet of workbook.sheets) {
    builder.addSheet(sheet.id, sheet.name);
  }

  return builder.build();
}

export { LayoutCursor } from "./layout-cursor";
export { RenderPlanBuilder } from "./render-plan-builder";
