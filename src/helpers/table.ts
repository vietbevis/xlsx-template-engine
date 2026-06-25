import type ExcelJS from 'exceljs';
import { f } from '../formula';
import { ReportEngineError } from '../errors';
import type {
  Block,
  CellContent,
  CellStyleDefinition,
  TableBorderDefinition,
  TableColumnNode,
  TableLeafColumn,
  TableSectionCell,
  TableSectionCellContext,
} from '../types';

export interface HeaderMatrixCell {
  title: string;
  style?: string | CellStyleDefinition;
  rowOffset: number;
  columnOffset: number;
  rowSpan: number;
  colSpan: number;
}

export function flattenColumns(columns: readonly TableColumnNode[]): TableLeafColumn[] {
  return columns.flatMap((column) =>
    column.children?.length ? flattenColumns(column.children) : [column as TableLeafColumn],
  );
}

export function calculateTableHeaderDepth(columns: readonly TableColumnNode[]): number {
  return Math.max(
    ...columns.map((column) => {
      if (column.children?.length) {
        return (column.childrenRowOffset ?? 1) + calculateTableHeaderDepth(column.children);
      }
      return 1;
    }),
  );
}

export function buildHeaderMatrix(columns: readonly TableColumnNode[], headerDepth: number): HeaderMatrixCell[] {
  const cells: HeaderMatrixCell[] = [];
  let columnOffset = 0;

  for (const column of columns) {
    columnOffset += appendHeaderCell(cells, column, 0, columnOffset, headerDepth);
  }

  return cells;
}

export function createTableColumnIdMap(columns: readonly TableLeafColumn[]): Map<string, number> {
  const idMap = new Map<string, number>();

  for (const [columnOffset, column] of columns.entries()) {
    if (!column.id) continue;

    const key = String(column.id);
    if (idMap.has(key)) throw new ReportEngineError(`Duplicate formula cell id "${key}".`);
    idMap.set(key, columnOffset);
  }

  return idMap;
}

export function resolveSectionCellColumnOffset(
  cell: TableSectionCell,
  cellIndex: number,
  tableColumnIdMap: Map<string, number>,
  occupiedColumns: Set<number>,
): number {
  if (cell.column !== undefined) return cell.column - 1;

  if (cell.columnId !== undefined) {
    const offset = tableColumnIdMap.get(cell.columnId);
    if (offset === undefined) {
      throw new ReportEngineError(`Table section row references unknown columnId "${cell.columnId}".`);
    }
    return offset;
  }

  let offset = cellIndex === 0 ? 0 : Math.max(...occupiedColumns) + 1;
  while (occupiedColumns.has(offset)) offset++;
  return offset;
}

export function resolveSectionCellColSpan(cell: TableSectionCell, columnOffset: number, tableWidth: number): number {
  const colSpan = cell.colSpan === 'remaining' ? tableWidth - columnOffset : (cell.colSpan ?? 1);

  if (columnOffset < 0 || columnOffset >= tableWidth || columnOffset + colSpan > tableWidth) {
    throw new ReportEngineError('Table section row cell range exceeds table width.');
  }

  return colSpan;
}

export function resolveSectionCellValue(cell: TableSectionCell, context: TableSectionCellContext): unknown {
  if (typeof cell.value === 'function') return cell.value(context);
  return cell.value ?? null;
}

export function createTableInlineStyle(border: TableBorderDefinition | undefined): CellStyleDefinition | undefined {
  if (!border) return undefined;
  return { border: typeof border === 'string' ? createAllSidesBorder(border) : border };
}

export function createSummaryFormula(columnId: string, summary: NonNullable<TableLeafColumn['summary']>): CellContent {
  if (summary && typeof summary === 'object') return summary;

  const range = f.range(columnId, columnId, { scope: 'allRows' });

  switch (summary) {
    case 'sum':
      return f`SUM(${range})`;
    case 'count':
      return f`COUNT(${range})`;
    case 'average':
      return f`AVERAGE(${range})`;
    default:
      return assertNeverSummary(summary);
  }
}

// ─── Private helpers ──────────────────────────────────────────────────────────

function appendHeaderCell(
  cells: HeaderMatrixCell[],
  column: TableColumnNode,
  rowOffset: number,
  columnOffset: number,
  headerDepth: number,
): number {
  const childColumns = column.children ?? [];
  const isParent = childColumns.length > 0;
  const childrenRowOffset = isParent ? (column.childrenRowOffset ?? 1) : 0;
  const colSpan = isParent ? countLeafColumns(childColumns) : 1;
  const rowSpan = isParent ? childrenRowOffset : headerDepth - rowOffset;

  cells.push({
    title: column.title,
    style: column.headerStyle ?? column.style,
    rowOffset,
    columnOffset,
    rowSpan,
    colSpan,
  });

  if (!isParent) return 1;

  let childColumnOffset = columnOffset;
  for (const childColumn of childColumns) {
    childColumnOffset += appendHeaderCell(
      cells,
      childColumn,
      rowOffset + childrenRowOffset,
      childColumnOffset,
      headerDepth,
    );
  }

  return colSpan;
}

function countLeafColumns(columns: readonly TableColumnNode[]): number {
  return columns.reduce(
    (total, column) => total + (column.children?.length ? countLeafColumns(column.children) : 1),
    0,
  );
}

function createAllSidesBorder(style: ExcelJS.BorderStyle): Partial<ExcelJS.Borders> {
  return { top: { style }, right: { style }, bottom: { style }, left: { style } };
}

function assertNeverSummary(value: never): never {
  throw new ReportEngineError(`Unsupported table summary "${String(value)}".`);
}
