import { ReportEngineError } from "../core/errors";
import type {
  Block,
  CellValue,
  SheetDefinition,
  StyleRegistry,
  WorkbookDefinition,
} from "../core/types";
import type { RenderPlanBuilder } from "./render-plan-builder";
import type { LayoutCursor } from "./layout-cursor";
import {
  interpolateCellValue,
  interpolateVariables,
  type VariableScope,
} from "./variable-engine";

export interface SheetContext {
  workbook: WorkbookDefinition;
  sheet: SheetDefinition;
  styles?: StyleRegistry;
  variables: VariableScope;
}

export type BlockCompiler<TBlock extends Block = Block> = (
  block: TBlock,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
) => void;

export type BlockCompilerRegistry = {
  [TType in Block["type"]]: BlockCompiler<Extract<Block, { type: TType }>>;
};

export const defaultBlockCompilerRegistry: BlockCompilerRegistry = {
  title(block, context, cursor, builder) {
    const variables = createBlockVariableScope(context, block);

    builder.addCell(context.sheet.id, {
      row: cursor.row,
      column: cursor.column,
      value: interpolateVariables(block.text, variables),
      style: block.style,
    });

    if (block.height !== undefined) {
      builder.setRowHeight(context.sheet.id, {
        row: cursor.row,
        height: block.height,
      });
    }

    cursor.advanceRows();
  },
  text(block, context, cursor, builder) {
    const variables = createBlockVariableScope(context, block);

    builder.addCell(context.sheet.id, {
      row: cursor.row,
      column: cursor.column,
      value: interpolateVariables(block.text, variables),
      style: block.style,
    });

    if (block.height !== undefined) {
      builder.setRowHeight(context.sheet.id, {
        row: cursor.row,
        height: block.height,
      });
    }

    cursor.advanceRows();
  },
  spacer(block, _context, cursor) {
    cursor.advanceRows(block.rows ?? 1);
  },
  grid(block, context, cursor, builder) {
    compileGridBlock(block, context, cursor, builder);
  },
  table(block, context, cursor, builder) {
    compileTableBlock(block, context, cursor, builder);
  },
};

export function compileBlock(
  block: Block,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
  registry: BlockCompilerRegistry = defaultBlockCompilerRegistry,
): void {
  const compiler = registry[block.type] as BlockCompiler | undefined;

  if (!compiler) {
    throw new ReportEngineError(`Unknown block type "${block.type}" in sheet "${context.sheet.id}".`);
  }

  compiler(block, context, cursor, builder);
}

function throwUnsupportedBlock(blockType: Block["type"]): never {
  throw new ReportEngineError(`Block type "${blockType}" is not supported by the compiler yet.`);
}

function compileGridBlock(
  block: Extract<Block, { type: "grid" }>,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
): void {
  const occupied = new Set<string>();
  let rowExtent = block.rows.length;
  const variables = createBlockVariableScope(context, block);

  for (const [rowOffset, gridRow] of block.rows.entries()) {
    const absoluteRow = cursor.row + rowOffset;

    if (gridRow.height !== undefined) {
      builder.setRowHeight(context.sheet.id, {
        row: absoluteRow,
        height: gridRow.height,
      });
    }

    let columnOffset = 0;

    for (const cell of gridRow.cells) {
      while (occupied.has(occupancyKey(rowOffset, columnOffset))) {
        columnOffset += 1;
      }

      const colSpan = cell.colSpan ?? 1;
      const rowSpan = cell.rowSpan ?? 1;
      assertGridCellDoesNotOverlap(occupied, rowOffset, columnOffset, rowSpan, colSpan);

      const absoluteColumn = cursor.column + columnOffset;

      builder.addCell(context.sheet.id, {
        row: absoluteRow,
        column: absoluteColumn,
        value: interpolateCellValue(cell.value, variables),
        style: cell.style,
      });

      if (cell.width !== undefined) {
        builder.setColumnWidth(context.sheet.id, {
          column: absoluteColumn,
          width: cell.width,
        });
      }

      if (rowSpan > 1 || colSpan > 1) {
        builder.addMerge(context.sheet.id, {
          startRow: absoluteRow,
          startColumn: absoluteColumn,
          endRow: absoluteRow + rowSpan - 1,
          endColumn: absoluteColumn + colSpan - 1,
        });
      }

      markGridCellOccupied(occupied, rowOffset, columnOffset, rowSpan, colSpan);
      rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
      columnOffset += colSpan;
    }
  }

  cursor.advanceRows(rowExtent);
}

function assertGridCellDoesNotOverlap(
  occupied: Set<string>,
  rowOffset: number,
  columnOffset: number,
  rowSpan: number,
  colSpan: number,
): void {
  for (let row = rowOffset; row < rowOffset + rowSpan; row += 1) {
    for (let column = columnOffset; column < columnOffset + colSpan; column += 1) {
      if (occupied.has(occupancyKey(row, column))) {
        throw new ReportEngineError("Grid cell merge ranges must not overlap.");
      }
    }
  }
}

function markGridCellOccupied(
  occupied: Set<string>,
  rowOffset: number,
  columnOffset: number,
  rowSpan: number,
  colSpan: number,
): void {
  for (let row = rowOffset; row < rowOffset + rowSpan; row += 1) {
    for (let column = columnOffset; column < columnOffset + colSpan; column += 1) {
      occupied.add(occupancyKey(row, column));
    }
  }
}

function occupancyKey(rowOffset: number, columnOffset: number): string {
  return `${rowOffset}:${columnOffset}`;
}

function compileTableBlock(
  block: Extract<Block, { type: "table" }>,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
): void {
  if (!Array.isArray(block.data)) {
    throw new ReportEngineError("Table async iterable data is not supported until streaming renderer phase 15.");
  }

  const variables = createBlockVariableScope(context, block);
  const headerDepth = calculateHeaderDepth(block.columns);
  const headerCells = buildHeaderMatrix(block.columns, headerDepth);
  const leafColumns = flattenLeafColumns(block.columns);

  for (const headerCell of headerCells) {
    builder.addCell(context.sheet.id, {
      row: cursor.row + headerCell.rowOffset,
      column: cursor.column + headerCell.columnOffset,
      value: interpolateVariables(headerCell.title, variables),
      style: block.headerStyle,
    });

    if (headerCell.rowSpan > 1 || headerCell.colSpan > 1) {
      builder.addMerge(context.sheet.id, {
        startRow: cursor.row + headerCell.rowOffset,
        startColumn: cursor.column + headerCell.columnOffset,
        endRow: cursor.row + headerCell.rowOffset + headerCell.rowSpan - 1,
        endColumn: cursor.column + headerCell.columnOffset + headerCell.colSpan - 1,
      });
    }
  }

  for (const [columnOffset, column] of leafColumns.entries()) {
    if (column.width !== undefined) {
      builder.setColumnWidth(context.sheet.id, {
        column: cursor.column + columnOffset,
        width: column.width,
      });
    }
  }

  for (const [rowOffset, rowData] of block.data.entries()) {
    const absoluteRow = cursor.row + headerDepth + rowOffset;

    for (const [columnOffset, column] of leafColumns.entries()) {
      const value = resolveTableCellValue(rowData, column);
      assertTableCellValue(value);

      builder.addCell(context.sheet.id, {
        row: absoluteRow,
        column: cursor.column + columnOffset,
        value: interpolateCellValue(value, variables),
        style: column.style ?? block.bodyStyle,
      });
    }
  }

  cursor.advanceRows(block.data.length + headerDepth);
}

function createBlockVariableScope(context: SheetContext, block: Block): VariableScope {
  return {
    workbook: context.variables.workbook,
    sheet: context.variables.sheet,
    block: block.context,
  };
}

function resolveTableCellValue(
  row: Record<string, unknown>,
  column: TableLeafColumn,
): unknown {
  if (column.accessor) {
    return column.accessor(row);
  }

  const value = column.key ? row[column.key] : null;
  return value === undefined ? null : value;
}

function assertTableCellValue(value: unknown): asserts value is CellValue {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    value instanceof Date
  ) {
    return;
  }

  throw new ReportEngineError("Table cell values must resolve to a supported cell value.");
}

interface HeaderMatrixCell {
  title: string;
  rowOffset: number;
  columnOffset: number;
  rowSpan: number;
  colSpan: number;
}

type TableColumnNode = Extract<Block, { type: "table" }>["columns"][number];
type TableLeafColumn = TableColumnNode & { children?: undefined };

function calculateHeaderDepth(columns: TableColumnNode[]): number {
  return Math.max(...columns.map((column) => {
    if (column.children && column.children.length > 0) {
      return 1 + calculateHeaderDepth(column.children);
    }

    return 1;
  }));
}

function flattenLeafColumns(columns: TableColumnNode[]): TableLeafColumn[] {
  return columns.flatMap((column) => {
    if (column.children && column.children.length > 0) {
      return flattenLeafColumns(column.children);
    }

    return [column as TableLeafColumn];
  });
}

function buildHeaderMatrix(
  columns: TableColumnNode[],
  headerDepth: number,
): HeaderMatrixCell[] {
  const cells: HeaderMatrixCell[] = [];
  let columnOffset = 0;

  for (const column of columns) {
    columnOffset += appendHeaderCell(cells, column, 0, columnOffset, headerDepth);
  }

  return cells;
}

function appendHeaderCell(
  cells: HeaderMatrixCell[],
  column: TableColumnNode,
  rowOffset: number,
  columnOffset: number,
  headerDepth: number,
): number {
  const childColumns = column.children ?? [];
  const isParent = childColumns.length > 0;
  const colSpan = isParent ? countLeafColumns(childColumns) : 1;
  const rowSpan = isParent ? 1 : headerDepth - rowOffset;

  cells.push({
    title: column.title,
    rowOffset,
    columnOffset,
    rowSpan,
    colSpan,
  });

  if (!isParent) {
    return 1;
  }

  let childColumnOffset = columnOffset;

  for (const childColumn of childColumns) {
    childColumnOffset += appendHeaderCell(
      cells,
      childColumn,
      rowOffset + 1,
      childColumnOffset,
      headerDepth,
    );
  }

  return colSpan;
}

function countLeafColumns(columns: TableColumnNode[]): number {
  return columns.reduce((total, column) => {
    if (column.children && column.children.length > 0) {
      return total + countLeafColumns(column.children);
    }

    return total + 1;
  }, 0);
}
