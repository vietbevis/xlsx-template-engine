import { AddressRegistry } from './address-registry';
import { compileBlock } from './block-compiler';
import type { Writable } from 'stream';
import { CompileError, FormulaError, ReportEngineError, ValidationError } from './errors';
import { cloneStylePart } from './helpers/style';
import { flattenColumns } from './helpers/table';
import ExcelJS from 'exceljs';
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
 * Compile `WorkbookDefinition` → `ExcelJS.Workbook` in a single pass.
 * Blocks are compiled directly into ExcelJS worksheets — no intermediate
 * representation.
 */
export function compileWorkbook(workbook: WorkbookDefinition, options: CompileWorkbookOptions = {}): ExcelJS.Workbook {
  validateWorkbookDefinition(workbook);

  const registry = new AddressRegistry();
  const workbookContext = { ...(workbook.context ?? {}), ...(options.context ?? {}) };

  const excelWorkbook = new ExcelJS.Workbook();
  if (workbook.metadata?.author) {
    excelWorkbook.creator = workbook.metadata.author;
  }

  const styleConfig = {
    defaultStyle: workbook.defaultStyle
      ? (cloneStylePart(workbook.defaultStyle) as typeof workbook.defaultStyle)
      : undefined,
    styles: workbook.styles,
  };

  for (const sheet of workbook.sheets) {
    const worksheet = excelWorkbook.addWorksheet(sheet.name, {
      views: sheet.freezePane
        ? [{ state: 'frozen' as const, xSplit: sheet.freezePane.columns ?? 0, ySplit: sheet.freezePane.rows ?? 0 }]
        : undefined,
    });

    const context = {
      workbook,
      sheet,
      worksheet,
      styleConfig,
      sheetColumnCount: measureSheetColumnCount(sheet),
      variables: { workbook: workbookContext, sheet: sheet.context },
      registry,
    };

    let row = 1;

    for (const [blockIndex, block] of sheet.blocks.entries()) {
      try {
        row = compileBlock(block, context, row);
      } catch (error) {
        throw normalizeCompileError(error, sheet.id, blockIndex);
      }
    }
  }

  return excelWorkbook;
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

// Removed cloneStyles as SheetWriter clones on assignment.

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
      // flattenColumns is also called inside TableBlockCompiler, but that
      // happens later per-row. Here we only traverse the column tree once
      // to get the width — cheap compared to data iteration.
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

export interface WorkbookRenderOptions {
  context?: RenderContext;
}

export interface WorkbookRenderer {
  writeFile(filePath: string): Promise<void>;
  writeBuffer(): Promise<Buffer>;
  writeStream(stream: Writable): Promise<void>;
}

export function renderWorkbook(workbook: WorkbookDefinition, options: CompileWorkbookOptions = {}): WorkbookRenderer {
  const excelWorkbook = compileWorkbook(workbook, options);

  return {
    writeFile: async (filePath: string) => {
      await excelWorkbook.xlsx.writeFile(filePath);
    },
    writeBuffer: async () => {
      const buffer = await excelWorkbook.xlsx.writeBuffer();
      return Buffer.from(buffer);
    },
    writeStream: async (stream: Writable) => {
      await excelWorkbook.xlsx.write(stream);
    },
  };
}
