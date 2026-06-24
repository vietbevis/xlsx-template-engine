import { ReportEngineError } from '../core/errors';
import { createSheetRegistry, type SheetRegistry } from '../core/sheet-registry';
import type {
  Block,
  FormulaDefinition,
  SheetDefinition,
  TableSectionRow,
  WorkbookDefinition,
} from '../core/types';
import { validateWorkbookDefinition } from '../core/validation';
import { compileBlock } from './block-compiler';
import { createFormulaKey, type CellAddress } from './formula-engine';
import { LayoutCursor } from './layout-cursor';
import type { RenderPlan } from './render-plan';
import { RenderPlanBuilder } from './render-plan-builder';
import type { RenderContext } from './variable-engine';

export interface CompileWorkbookOptions {
  context?: RenderContext;
}

export function compileWorkbookToRenderPlan(
  workbook: WorkbookDefinition,
  options: CompileWorkbookOptions = {},
): RenderPlan {
  validateWorkbookDefinition(workbook);
  const sheetRegistry = createSheetRegistry(workbook.sheets);
  collectFormulaDependencies(workbook, sheetRegistry);
  const formulaKeys = collectWorkbookFormulaKeys(workbook);

  const workbookContext = {
    ...(workbook.context ?? {}),
    ...(options.context ?? {}),
  };
  const builder = new RenderPlanBuilder(workbook.metadata, workbook.defaultStyle, workbook.styles);

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
      formulaKeys,
    };
    const cursor = new LayoutCursor();

    for (const block of sheet.blocks) {
      compileBlock(block, context, cursor, builder);
    }
  }

  return builder.build();
}

export type FormulaDependencyGraph = ReadonlyMap<string, readonly string[]>;

export function collectFormulaDependencies(
  workbook: WorkbookDefinition,
  sheetRegistry: SheetRegistry = createSheetRegistry(workbook.sheets),
): FormulaDependencyGraph {
  const dependencies = new Map<string, Set<string>>();

  for (const sheet of workbook.sheets) {
    const sheetDependencies = new Set<string>();

    for (const formula of collectSheetFormulas(sheet)) {
      collectFormulaSheetIds(formula, sheet.id, sheetRegistry, sheetDependencies);
    }

    dependencies.set(sheet.id, sheetDependencies);
  }

  return new Map(
    Array.from(dependencies.entries()).map(([sheetId, sheetDependencies]) => [
      sheetId,
      Array.from(sheetDependencies),
    ]),
  );
}

function collectWorkbookFormulaKeys(workbook: WorkbookDefinition): Map<string, CellAddress> {
  const formulaKeys = new Map<string, CellAddress>();

  for (const sheet of workbook.sheets) {
    const cursor = new LayoutCursor();

    for (const block of sheet.blocks) {
      switch (block.type) {
        case 'title':
        case 'text':
          cursor.advanceRows();
          break;
        case 'spacer':
          cursor.advanceRows(block.rows ?? 1);
          break;
        case 'grid':
          cursor.advanceRows(collectGridFormulaKeys(formulaKeys, sheet, block, cursor));
          break;
        case 'table':
          if (!Array.isArray(block.data)) {
            throw new ReportEngineError(
              'Table async iterable data is not supported until streaming renderer phase 15.',
            );
          }

          cursor.advanceRows(collectTableFormulaKeys(formulaKeys, sheet, block, cursor));
          break;
        default:
          assertNever(block);
      }
    }
  }

  return formulaKeys;
}

function collectGridFormulaKeys(
  formulaKeys: Map<string, CellAddress>,
  sheet: SheetDefinition,
  block: Extract<Block, { type: 'grid' }>,
  cursor: LayoutCursor,
): number {
  const occupied = new Set<string>();
  let rowExtent = block.rows.length;

  for (const [rowOffset, gridRow] of block.rows.entries()) {
    let columnOffset = 0;

    for (const cell of gridRow.cells) {
      while (occupied.has(gridOccupancyKey(rowOffset, columnOffset))) {
        columnOffset += 1;
      }

      const rowSpan = cell.rowSpan ?? 1;
      const colSpan = cell.colSpan ?? 1;
      const row = cursor.row + rowOffset;
      const column = cursor.column + columnOffset;

      if (cell.key) {
        registerWorkbookFormulaKey(formulaKeys, sheet, cell.key, row, column);
      }

      markGridOccupied(occupied, rowOffset, columnOffset, rowSpan, colSpan);
      rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
      columnOffset += colSpan;
    }
  }

  return rowExtent;
}

function registerWorkbookFormulaKey(
  formulaKeys: Map<string, CellAddress>,
  sheet: SheetDefinition,
  key: string,
  row: number,
  column: number,
): void {
  const registryKey = createFormulaKey(sheet.id, key);

  if (formulaKeys.has(registryKey)) {
    throw new ReportEngineError(`Duplicate formula cell key "${key}" in sheet "${sheet.id}".`);
  }

  formulaKeys.set(registryKey, {
    row,
    column,
    sheetId: sheet.id,
    sheetName: sheet.name,
  });
}

function collectSheetFormulas(sheet: SheetDefinition): FormulaDefinition[] {
  const formulas: FormulaDefinition[] = [];

  for (const block of sheet.blocks) {
    if (block.type === 'grid') {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (isFormulaObject(cell.value)) {
            formulas.push(cell.value);
          }
        }
      }
    }

    if (block.type === 'table' && Array.isArray(block.data)) {
      for (const column of flattenColumns(block.columns)) {
        for (const row of block.data) {
          if (isTableSectionRow(row)) {
            continue;
          }

          if (!column.accessor) {
            const value = column.key ? row[column.key] : undefined;

            if (isFormulaObject(value)) {
              formulas.push(value);
            }

            continue;
          }

          const value = column.accessor(row);

          if (isFormulaObject(value)) {
            formulas.push(value);
          }
        }
      }

      let currentRows: Record<string, unknown>[] = [];
      const allRows = collectAllTableRows(block);

      for (const [dataIndex, item] of block.data.entries()) {
        if (isTableSectionRow(item)) {
          for (const cell of item.cells) {
            if (typeof cell.value === 'function') {
              const value = cell.value({
                rows: currentRows,
                allRows,
                dataIndex,
                rowIndex: 1,
              });

              if (isFormulaObject(value)) {
                formulas.push(value);
              }
            } else if (isFormulaObject(cell.value)) {
              formulas.push(cell.value);
            }
          }

          if (item.resetRows) {
            currentRows = [];
          }

          continue;
        }

        currentRows.push(item);
      }
    }
  }

  return formulas;
}

function calculateTableRowExtent(block: Extract<Block, { type: 'table' }>): number {
  const titleRows = block.titleRows?.length ?? 0;
  const headerRows = calculateTableHeaderDepth(block.columns);
  const dataRows = Array.isArray(block.data) ? block.data.length : 0;

  return titleRows + headerRows + dataRows;
}

function collectTableFormulaKeys(
  formulaKeys: Map<string, CellAddress>,
  sheet: SheetDefinition,
  block: Extract<Block, { type: 'table' }>,
  cursor: LayoutCursor,
): number {
  const headerDepth = calculateTableHeaderDepth(block.columns);
  const columnKeyMap = createTableColumnKeyMap(flattenColumns(block.columns));
  let rowOffset = (block.titleRows?.length ?? 0) + headerDepth;

  const data = block.data;

  if (!Array.isArray(data)) {
    throw new ReportEngineError(
      'Table async iterable data is not supported until streaming renderer phase 15.',
    );
  }

  for (const item of data) {
    if (isTableSectionRow(item)) {
      collectTableSectionFormulaKeys(
        formulaKeys,
        sheet,
        item,
        cursor.row + rowOffset,
        cursor.column,
        columnKeyMap,
      );
    }

    rowOffset += 1;
  }

  return rowOffset;
}

function collectTableSectionFormulaKeys(
  formulaKeys: Map<string, CellAddress>,
  sheet: SheetDefinition,
  sectionRow: TableSectionRow<Record<string, unknown>>,
  row: number,
  firstColumn: number,
  columnKeyMap: Map<string, number>,
): void {
  const occupiedColumns = new Set<number>();

  for (const [cellIndex, cell] of sectionRow.cells.entries()) {
    const columnOffset = resolveTableSectionCellColumnOffset(
      cell,
      cellIndex,
      columnKeyMap,
      occupiedColumns,
    );
    const colSpan = cell.colSpan === 'remaining' ? 1 : (cell.colSpan ?? 1);

    if (cell.key) {
      registerWorkbookFormulaKey(formulaKeys, sheet, cell.key, row, firstColumn + columnOffset);
    }

    for (let offset = columnOffset; offset < columnOffset + colSpan; offset += 1) {
      occupiedColumns.add(offset);
    }
  }
}

function resolveTableSectionCellColumnOffset(
  cell: TableSectionRow<Record<string, unknown>>['cells'][number],
  cellIndex: number,
  columnKeyMap: Map<string, number>,
  occupiedColumns: Set<number>,
): number {
  if (cell.column !== undefined) {
    return cell.column - 1;
  }

  if (cell.columnKey !== undefined) {
    const offset = columnKeyMap.get(cell.columnKey);

    if (offset === undefined) {
      throw new ReportEngineError(
        `Table section row references unknown columnKey "${cell.columnKey}".`,
      );
    }

    return offset;
  }

  let offset = cellIndex === 0 ? 0 : Math.max(...occupiedColumns) + 1;

  while (occupiedColumns.has(offset)) {
    offset += 1;
  }

  return offset;
}

function collectAllTableRows(block: Extract<Block, { type: 'table' }>): Record<string, unknown>[] {
  if (!Array.isArray(block.data)) {
    return [];
  }

  return block.data.filter((item) => !isTableSectionRow(item)) as Record<string, unknown>[];
}

function createTableColumnKeyMap(columns: TableLeafColumn[]): Map<string, number> {
  const keyMap = new Map<string, number>();

  for (const [columnOffset, column] of columns.entries()) {
    if (!column.key) {
      continue;
    }

    if (keyMap.has(String(column.key))) {
      throw new ReportEngineError(`Duplicate formula cell key "${String(column.key)}".`);
    }

    keyMap.set(String(column.key), columnOffset);
  }

  return keyMap;
}

function isTableSectionRow(value: unknown): value is TableSectionRow<Record<string, unknown>> {
  return typeof value === 'object' && value !== null && 'type' in value && value.type === 'section';
}

function collectFormulaSheetIds(
  formula: FormulaDefinition,
  currentSheetId: string,
  sheetRegistry: SheetRegistry,
  dependencies: Set<string>,
): void {
  if ('sheetId' in formula && formula.sheetId !== undefined) {
    if (!sheetRegistry.has(formula.sheetId)) {
      throw new ReportEngineError(`Formula references unknown sheetId "${formula.sheetId}".`);
    }

    if (formula.sheetId !== currentSheetId) {
      dependencies.add(formula.sheetId);
    }
  }

  switch (formula.type) {
    case 'raw':
    case 'literal':
    case 'ref':
    case 'range':
      return;
    case 'sum':
      formula.values?.forEach((value) =>
        collectFormulaSheetIds(value, currentSheetId, sheetRegistry, dependencies),
      );
      return;
    case 'round':
      collectFormulaSheetIds(formula.value, currentSheetId, sheetRegistry, dependencies);
      return;
    case 'if':
      collectFormulaSheetIds(formula.condition, currentSheetId, sheetRegistry, dependencies);
      collectFormulaSheetIds(formula.whenTrue, currentSheetId, sheetRegistry, dependencies);
      collectFormulaSheetIds(formula.whenFalse, currentSheetId, sheetRegistry, dependencies);
      return;
    case 'call':
      formula.args.forEach((arg) =>
        collectFormulaSheetIds(arg, currentSheetId, sheetRegistry, dependencies),
      );
      return;
    case 'binary':
      collectFormulaSheetIds(formula.left, currentSheetId, sheetRegistry, dependencies);
      collectFormulaSheetIds(formula.right, currentSheetId, sheetRegistry, dependencies);
      return;
    default:
      assertNever(formula);
  }
}

function isFormulaObject(value: unknown): value is FormulaDefinition {
  return typeof value === 'object' && value !== null && 'type' in value;
}

type TableColumnNode = Extract<Block, { type: 'table' }>['columns'][number];
type TableLeafColumn = TableColumnNode & { children?: undefined };

function flattenColumns(columns: readonly TableColumnNode[]): TableLeafColumn[] {
  return columns.flatMap((column) => {
    if (column.children && column.children.length > 0) {
      return flattenColumns(column.children);
    }

    return [column as TableLeafColumn];
  });
}

function calculateTableHeaderDepth(columns: readonly TableColumnNode[]): number {
  return Math.max(
    ...columns.map((column) => {
      if (column.children && column.children.length > 0) {
        return 1 + calculateTableHeaderDepth(column.children);
      }

      return 1;
    }),
  );
}

function markGridOccupied(
  occupied: Set<string>,
  rowOffset: number,
  columnOffset: number,
  rowSpan: number,
  colSpan: number,
): void {
  for (let row = rowOffset; row < rowOffset + rowSpan; row += 1) {
    for (let column = columnOffset; column < columnOffset + colSpan; column += 1) {
      occupied.add(gridOccupancyKey(row, column));
    }
  }
}

function gridOccupancyKey(rowOffset: number, columnOffset: number): string {
  return `${rowOffset}:${columnOffset}`;
}

function assertNever(value: never): never {
  throw new ReportEngineError(`Unsupported block type "${(value as Block).type}".`);
}

export { compileBlock, defaultBlockCompilerRegistry } from './block-compiler';
export {
  compileCellContent,
  compileFormula,
  createFormulaCompileContext,
  createFormulaKey,
  formatCellAddress,
  formatCellReference,
  isFormulaDefinition,
} from './formula-engine';
export { LayoutCursor } from './layout-cursor';
export { assertMergeDoesNotOverlap, normalizeMergeRange } from './merge-engine';
export { RenderPlanBuilder } from './render-plan-builder';
export { interpolateCellValue, interpolateVariables, resolvePath } from './variable-engine';

export type { BlockCompiler, BlockCompilerRegistry, SheetContext } from './block-compiler';
export type {
  CellAddress,
  FormulaCompileContext,
  FormulaCompileContextOptions,
} from './formula-engine';
export type { MergeRange } from './merge-engine';
export type { RenderContext, VariableScope } from './variable-engine';
