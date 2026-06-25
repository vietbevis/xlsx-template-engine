import { AddressRegistry } from './address-registry';
import { compileBlock } from './block-compiler';
import { cloneStylePart } from './helpers/style';
import { createSheetColumnCounts } from './helpers/workbook';
import type { RenderPlan } from './render-plan';
import { SheetWriter } from './sheet-writer';
import type { WorkbookDefinition } from './types';
import { validateWorkbookDefinition } from './validation';
import type { RenderContext } from './variable-engine';

export interface CompileWorkbookOptions {
  context?: RenderContext;
}

/**
 * Compile `WorkbookDefinition` → `RenderPlan` trong **single-pass**.
 *
 * Loại bỏ hoàn toàn:
 * - `createSheetColumnCounts` pre-pass (tích hợp vào context)
 * - `collectWorkbookFormulaIds` pre-pass (thay bằng AddressRegistry ghi inline)
 * - `RenderPlanBuilder` (thay bằng `SheetWriter` per-sheet)
 * - `LayoutCursor` mutable (thay bằng `row` number thuần)
 */
export function compileWorkbookToRenderPlan(
  workbook: WorkbookDefinition,
  options: CompileWorkbookOptions = {},
): RenderPlan {
  validateWorkbookDefinition(workbook);

  const sheetColumnCounts = createSheetColumnCounts(workbook);
  const registry = new AddressRegistry();
  const workbookContext = { ...(workbook.context ?? {}), ...(options.context ?? {}) };

  const sheets = workbook.sheets.map((sheet) => {
    const writer = new SheetWriter(
      sheet.id,
      sheet.name,
      sheet.freezePane
        ? [{ state: 'frozen', xSplit: sheet.freezePane.columns ?? 0, ySplit: sheet.freezePane.rows ?? 0 }]
        : undefined,
    );

    const context = {
      workbook,
      sheet,
      sheetColumnCount: sheetColumnCounts.get(sheet.id) ?? 1,
      variables: { workbook: workbookContext, sheet: sheet.context },
      registry,
    };

    let row = 1;

    for (const block of sheet.blocks) {
      row = compileBlock(block, context, writer, row);
    }

    return writer.finish();
  });

  return {
    metadata: workbook.metadata
      ? {
          ...workbook.metadata,
          keywords: workbook.metadata.keywords ? [...workbook.metadata.keywords] : undefined,
        }
      : undefined,
    defaultStyle: workbook.defaultStyle
      ? (cloneStylePart(workbook.defaultStyle) as typeof workbook.defaultStyle)
      : undefined,
    styles: workbook.styles ? resolveStyles(workbook.styles) : undefined,
    sheets,
  };
}

function resolveStyles(styles: NonNullable<WorkbookDefinition['styles']>): NonNullable<WorkbookDefinition['styles']> {
  return Object.fromEntries(
    Object.entries(styles).map(([name, style]) => [name, cloneStylePart(style) as typeof style]),
  );
}
