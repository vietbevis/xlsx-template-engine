import ExcelJS from 'exceljs';
import type { Writable } from 'stream';
import { AddressRegistry } from '../formula/address-registry';
import { CompileError, FormulaError, ReportEngineError, ValidationError } from '../errors';
import { cloneStylePart } from '../styles/style-resolver';
import { flattenColumns } from '../utils/table-utils';
import type { Block, SheetDefinition, WorkbookDefinition } from '../types';
import { validateWorkbookDefinition } from '../validation';
import { compileBlock } from './block-dispatcher';

// ─── Public interfaces ────────────────────────────────────────────────────────

export interface CompileWorkbookOptions {}

export interface WorkbookRenderOptions {}

export interface WorkbookRenderer {
  writeFile(filePath: string): Promise<void>;
  writeBuffer(): Promise<Buffer>;
  writeStream(stream: Writable): Promise<void>;
}

// ─── WorkbookCompiler ─────────────────────────────────────────────────────────

/**
 * Compile `WorkbookDefinition` → `ExcelJS.Workbook` in a single pass.
 * Blocks are compiled directly into ExcelJS worksheets — no intermediate
 * representation.
 */
export class WorkbookCompiler {
  constructor(
    private readonly workbook: WorkbookDefinition,
    private readonly options: CompileWorkbookOptions = {},
  ) {}

  compile(): ExcelJS.Workbook {
    validateWorkbookDefinition(this.workbook);

    const registry = new AddressRegistry();

    const excelWorkbook = new ExcelJS.Workbook();
    if (this.workbook.metadata?.author) {
      excelWorkbook.creator = this.workbook.metadata.author;
    }

    const styleConfig = {
      defaultStyle: this.workbook.defaultStyle
        ? (cloneStylePart(this.workbook.defaultStyle) as typeof this.workbook.defaultStyle)
        : undefined,
      styles: this.workbook.styles,
    };

    for (const sheet of this.workbook.sheets) {
      const worksheet = excelWorkbook.addWorksheet(sheet.name);

      const context = {
        workbook: this.workbook,
        sheet,
        worksheet,
        styleConfig,
        sheetColumnCount: WorkbookCompiler.measureSheetColumnCount(sheet),
        registry,
      };

      let row = 1;

      for (const [blockIndex, block] of sheet.blocks.entries()) {
        try {
          row = compileBlock(block, context, row);
        } catch (error) {
          throw WorkbookCompiler.normalizeCompileError(error, sheet.id, blockIndex);
        }
      }
    }

    return excelWorkbook;
  }

  static render(workbook: WorkbookDefinition, options: CompileWorkbookOptions = {}): WorkbookRenderer {
    const compiler = new WorkbookCompiler(workbook, options);
    const excelWorkbook = compiler.compile();

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

  // ─── Private helpers ──────────────────────────────────────────────────────

  private static normalizeCompileError(error: unknown, sheetId: string, blockIndex: number): Error {
    if (error instanceof ValidationError || error instanceof FormulaError || error instanceof CompileError) {
      return error;
    }
    if (error instanceof ReportEngineError) {
      return new CompileError(error.message, { sheetId, blockIndex });
    }
    return error instanceof Error ? error : new CompileError(String(error), { sheetId, blockIndex });
  }

  private static measureSheetColumnCount(sheet: SheetDefinition): number {
    return Math.max(1, ...sheet.blocks.map(WorkbookCompiler.measureBlockColumnCount));
  }

  private static measureBlockColumnCount(block: Block): number {
    switch (block.type) {
      case 'grid':
        return Math.max(
          1,
          ...block.rows.map((row) =>
            row.cells.reduce((width, cell) => width + (cell.colSpan === 'remaining' ? 1 : (cell.colSpan ?? 1)), 0),
          ),
        );
      case 'table':
        return flattenColumns(block.columns).length;
      default:
        throw new CompileError(`Unsupported block type "${(block as Block).type}".`);
    }
  }
}

// ─── Convenience functions ────────────────────────────────────────────────────

export function compileWorkbook(workbook: WorkbookDefinition, options: CompileWorkbookOptions = {}): ExcelJS.Workbook {
  return new WorkbookCompiler(workbook, options).compile();
}

export function renderWorkbook(workbook: WorkbookDefinition, options: CompileWorkbookOptions = {}): WorkbookRenderer {
  return WorkbookCompiler.render(workbook, options);
}
