import { AddressRegistry } from './address-registry';
import { compileBlock } from './block-compiler';
import { CompileError, FormulaError, ReportEngineError, ValidationError } from './errors';
import { cloneStylePart } from './helpers/style';
import { flattenColumns } from './helpers/table';
import type { RenderPlan } from './render-plan';
import { SheetWriter } from './sheet-writer';
import type { Block, GridRow, SheetDefinition, WorkbookDefinition } from './types';
import { validateWorkbookDefinition } from './validation';
import type { RenderContext } from './variable-engine';

export interface CompileWorkbookOptions {
  /**
   * Runtime context is merged over `workbook.context`, so values passed here
   * intentionally override same-named definition defaults.
   */
  context?: RenderContext;
}

/**
 * Compile `WorkbookDefinition` → `RenderPlan` trong **single-pass**.
 *
 * Loại bỏ hoàn toàn:
 * - workbook-level column-count pre-pass (sheet width is measured per sheet)
 * - `collectWorkbookFormulaIds` pre-pass (thay bằng AddressRegistry ghi inline)
 * - `RenderPlanBuilder` (thay bằng `SheetWriter` per-sheet)
 * - `LayoutCursor` mutable (thay bằng `row` number thuần)
 */
export function compileWorkbookToRenderPlan(
  workbook: WorkbookDefinition,
  options: CompileWorkbookOptions = {},
): RenderPlan {
  validateWorkbookDefinition(workbook);

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
      sheetColumnCount: measureSheetColumnCount(sheet),
      variables: { workbook: workbookContext, sheet: sheet.context },
      registry,
    };

    let row = 1;

    for (const [blockIndex, block] of sheet.blocks.entries()) {
      try {
        row = compileBlock(block, context, writer, row);
      } catch (error) {
        throw normalizeCompileError(error, sheet.id, blockIndex);
      }
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

function normalizeCompileError(error: unknown, sheetId: string, blockIndex: number): Error {
  if (error instanceof ValidationError || error instanceof FormulaError || error instanceof CompileError) {
    return error;
  }

  if (error instanceof ReportEngineError) {
    return new CompileError(error.message, { sheetId, blockIndex });
  }

  return error instanceof Error ? error : new CompileError(String(error), { sheetId, blockIndex });
}

function resolveStyles(styles: NonNullable<WorkbookDefinition['styles']>): NonNullable<WorkbookDefinition['styles']> {
  return Object.fromEntries(
    Object.entries(styles).map(([name, style]) => [name, cloneStylePart(style) as typeof style]),
  );
}

function measureSheetColumnCount(sheet: SheetDefinition): number {
  return Math.max(1, ...sheet.blocks.map(measureBlockColumnCount));
}

function measureBlockColumnCount(block: Block): number {
  switch (block.type) {
    case 'title':
    case 'text':
      return block.colSpan === 'remaining' ? 1 : (block.colSpan ?? 1);
    case 'spacer':
    case 'divider':
      return 1;
    case 'grid':
      return Math.max(1, ...block.rows.map(measureGridRowColumnCount));
    case 'table':
    case 'table-groups':
      return flattenColumns(block.columns).length;
    default:
      return assertNeverBlock(block);
  }
}

function measureGridRowColumnCount(row: GridRow): number {
  return row.cells.reduce((width, cell) => width + (cell.colSpan === 'remaining' ? 1 : (cell.colSpan ?? 1)), 0);
}

function assertNeverBlock(block: never): never {
  throw new CompileError(`Unsupported block type "${(block as Block).type}".`);
}
