import type ExcelJS from 'exceljs';
import { AddressRegistry } from '../formula/address-registry';
import { ReportEngineError } from '../errors';
import type { Block, CellStyleDefinition, SheetDefinition, StyleRegistry, WorkbookDefinition } from '../types';
import { compileGridBlock } from './grid-block-compiler';
import { compileTableBlock } from './table-block-compiler';

export interface CompileContext {
  readonly workbook: WorkbookDefinition;
  readonly sheet: SheetDefinition;
  readonly sheetColumnCount: number;
  readonly registry: AddressRegistry;
  readonly worksheet: ExcelJS.Worksheet;
  readonly styleConfig: {
    defaultStyle?: CellStyleDefinition;
    styles?: StyleRegistry;
  };
}

export function compileBlock(block: Block, context: CompileContext, startRow: number): number {
  switch (block.type) {
    case 'grid':
      return compileGridBlock(block, context, startRow);
    case 'table':
      return compileTableBlock(block, context, startRow);
    default:
      throw new ReportEngineError(`Unknown block type "${(block as Block).type}" in sheet "${context.sheet.id}".`);
  }
}
