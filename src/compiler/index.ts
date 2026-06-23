import type { WorkbookDefinition } from "../core/types";
import { validateWorkbookDefinition } from "../core/validation";
import { createSheetRegistry } from "../core/sheet-registry";
import type { RenderPlan } from "./render-plan";
import { RenderPlanBuilder } from "./render-plan-builder";
import { LayoutCursor } from "./layout-cursor";
import { compileBlock } from "./block-compiler";
import type { RenderContext } from "./variable-engine";

export interface CompileWorkbookOptions {
  context?: RenderContext;
}

export function compileWorkbookToRenderPlan(
  workbook: WorkbookDefinition,
  options: CompileWorkbookOptions = {},
): RenderPlan {
  validateWorkbookDefinition(workbook);
  createSheetRegistry(workbook.sheets);

  const workbookContext = {
    ...(workbook.context ?? {}),
    ...(options.context ?? {}),
  };
  const builder = new RenderPlanBuilder(workbook.metadata, workbook.styles);

  for (const sheet of workbook.sheets) {
    builder.addSheet(sheet.id, sheet.name);
  }

  for (const sheet of workbook.sheets) {
    const context = {
      workbook,
      sheet,
      styles: workbook.styles,
      variables: {
        workbook: workbookContext,
        sheet: sheet.context,
      },
    };
    const cursor = new LayoutCursor();

    for (const block of sheet.blocks) {
      compileBlock(block, context, cursor, builder);
    }
  }

  return builder.build();
}

export { LayoutCursor } from "./layout-cursor";
export { RenderPlanBuilder } from "./render-plan-builder";
export {
  compileBlock,
  defaultBlockCompilerRegistry,
} from "./block-compiler";
export {
  assertMergeDoesNotOverlap,
  normalizeMergeRange,
} from "./merge-engine";
export {
  interpolateCellValue,
  interpolateVariables,
  resolvePath,
} from "./variable-engine";

export type {
  BlockCompiler,
  BlockCompilerRegistry,
  SheetContext,
} from "./block-compiler";
export type {
  MergeRange,
} from "./merge-engine";
export type {
  RenderContext,
  VariableScope,
} from "./variable-engine";
