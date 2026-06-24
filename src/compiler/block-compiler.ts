import type ExcelJS from 'exceljs';
import { ReportEngineError } from '../core/errors';
import type {
  Block,
  CellContent,
  CellStyleDefinition,
  FormulaRangeScope,
  GridCell,
  SheetDefinition,
  StyleRegistry,
  StyleValue,
  TableBorderDefinition,
  TableFooterRow,
  TableSectionCell,
  TableSectionCellContext,
  TableSectionRow,
  TableTitleRow,
  WorkbookDefinition,
} from '../core/types';
import {
  type CellAddress,
  compileCellContent,
  createFormulaCompileContext,
  createFormulaId,
  formatCellAddress,
  formatCellReference,
  type FormulaCompileContext,
  isFormulaDefinition,
} from './formula-engine';
import type { LayoutCursor } from './layout-cursor';
import type { RenderPlanBuilder } from './render-plan-builder';
import { interpolateCellValue, interpolateVariables, type VariableScope } from './variable-engine';

export interface SheetContext {
  workbook: WorkbookDefinition;
  sheet: SheetDefinition;
  styles?: StyleRegistry;
  variables: VariableScope;
  formulaIds?: Map<string, CellAddress>;
  namedRangeNames?: ReadonlySet<string>;
}

export type BlockCompiler<TBlock extends Block = Block> = (
  block: TBlock,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
) => void;

export type BlockCompilerRegistry = {
  [TType in Block['type']]: BlockCompiler<Extract<Block, { type: TType }>>;
};

const defaultBlockCompilerRegistry: BlockCompilerRegistry = {
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
  divider(block, context, cursor, builder) {
    const rows = block.rows ?? 1;

    for (let rowOffset = 0; rowOffset < rows; rowOffset += 1) {
      builder.addCell(context.sheet.id, {
        row: cursor.row + rowOffset,
        column: cursor.column,
        value: '',
        style: block.style,
      });
    }

    cursor.advanceRows(rows);
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

function compileGridBlock(
  block: Extract<Block, { type: 'grid' }>,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
): void {
  const occupied = new Set<string>();
  const placements: GridCellPlacement[] = [];
  const variables = createBlockVariableScope(context, block);
  let rowExtent = block.rows.length;

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
      placements.push({
        cell,
        rowOffset,
        columnOffset,
        rowSpan,
        colSpan,
        row: absoluteRow,
        column: absoluteColumn,
      });

      markGridCellOccupied(occupied, rowOffset, columnOffset, rowSpan, colSpan);
      rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
      columnOffset += colSpan;
    }
  }

  const formulaContext = createFormulaCompileContext(
    context.formulaIds ?? createGridCellIdMap(placements, context.sheet),
    { currentSheetId: context.sheet.id, namedRanges: context.namedRangeNames },
  );

  for (const placement of placements) {
    const compiledContent = compileCellContent(interpolateCellValue(placement.cell.value, variables), formulaContext);

    builder.addCell(context.sheet.id, {
      row: placement.row,
      column: placement.column,
      ...compiledContent,
      formulaResult: isFormulaDefinition(placement.cell.value) ? placement.cell.formulaResult : undefined,
      style: resolveCellStyle(
        placement.cell.style,
        placement.cell.styleResolver?.(compiledContent.value),
        context,
        `grid cell "${placement.cell.id ?? `${placement.row}:${placement.column}`}"`,
      ),
    });

    if (placement.cell.width !== undefined) {
      builder.setColumnWidth(context.sheet.id, {
        column: placement.column,
        width: placement.cell.width,
      });
    }

    if (placement.rowSpan > 1 || placement.colSpan > 1) {
      builder.addMerge(context.sheet.id, {
        startRow: placement.row,
        startColumn: placement.column,
        endRow: placement.row + placement.rowSpan - 1,
        endColumn: placement.column + placement.colSpan - 1,
      });
    }
  }

  cursor.advanceRows(rowExtent);
}

interface GridCellPlacement {
  cell: GridCell;
  rowOffset: number;
  columnOffset: number;
  rowSpan: number;
  colSpan: number;
  row: number;
  column: number;
}

function createGridCellIdMap(placements: GridCellPlacement[], sheet: SheetDefinition): Map<string, CellAddress> {
  const idMap = new Map<string, CellAddress>();

  for (const placement of placements) {
    if (!placement.cell.id) {
      continue;
    }

    registerCellId(idMap, createFormulaId(sheet.id, placement.cell.id), {
      row: placement.row,
      column: placement.column,
      sheetId: sheet.id,
      sheetName: sheet.name,
    });
  }

  return idMap;
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
        throw new ReportEngineError('Grid cell merge ranges must not overlap.');
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
  block: Extract<Block, { type: 'table' }>,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
): void {
  if (block.data !== undefined && !Array.isArray(block.data)) {
    throw new ReportEngineError('AsyncIterable table data is not yet supported. Use an array data source.');
  }

  const variables = createBlockVariableScope(context, block);
  const headerDepth = calculateHeaderDepth(block.columns);
  const headerCells = buildHeaderMatrix(block.columns, headerDepth);
  const leafColumns = flattenLeafColumns(block.columns);
  const tableColumnIdMap = createTableColumnIdMap(leafColumns);
  const tableWidth = leafColumns.length;
  const allRows = collectTableDataRows(block);
  const tableInlineStyle = createTableInlineStyle(block.border);
  const footerRows = resolveTableFooterRows(block, leafColumns);

  for (const titleRow of block.titleRows ?? []) {
    compileTableTitleRow(titleRow, context, cursor, builder, tableWidth, variables);
  }

  for (const [rowOffset, height] of (block.headerRowHeights ?? []).entries()) {
    builder.setRowHeight(context.sheet.id, {
      row: cursor.row + rowOffset,
      height,
    });
  }

  for (const headerCell of headerCells) {
    builder.addCell(context.sheet.id, {
      row: cursor.row + headerCell.rowOffset,
      column: cursor.column + headerCell.columnOffset,
      value: interpolateVariables(headerCell.title, variables),
      style: headerCell.style ?? block.headerStyle,
      inlineStyle: tableInlineStyle,
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

    if (column.hidden !== undefined) {
      builder.setColumnHidden(context.sheet.id, {
        column: cursor.column + columnOffset,
        hidden: column.hidden,
      });
    }
  }

  cursor.advanceRows(headerDepth);

  if (!Array.isArray(block.data)) {
    return;
  }

  const allRenderedRows: RenderedTableDataRow[] = [];
  let currentRenderedRows: RenderedTableDataRow[] = [];

  for (const [dataIndex, item] of block.data.entries()) {
    if (isTableSectionRow(item)) {
      compileTableSectionRow(item, {
        allRows,
        allRenderedRows,
        builder,
        context,
        currentRenderedRows,
        cursor,
        dataIndex,
        tableColumnIdMap,
        tableWidth,
        tableInlineStyle,
        variables,
      });

      if (item.resetRows) {
        currentRenderedRows = [];
      }

      continue;
    }

    const rowIndex = cursor.row;
    compileTableDataRow(
      item as Record<string, unknown>,
      dataIndex,
      context,
      cursor,
      builder,
      leafColumns,
      tableColumnIdMap,
      variables,
      block.bodyStyle,
      dataIndex % 2 === 0 ? block.oddRowStyle : block.evenRowStyle,
      tableInlineStyle,
      block.bodyRowHeight,
      block.rowHidden?.(item as Record<string, unknown>, dataIndex),
    );
    const renderedRow = {
      data: item as Record<string, unknown>,
      row: rowIndex,
    };
    currentRenderedRows.push(renderedRow);
    allRenderedRows.push(renderedRow);
  }

  for (const footerRow of footerRows) {
    compileTableFooterRow(footerRow, {
      allRows,
      allRenderedRows,
      builder,
      context,
      currentRenderedRows: allRenderedRows,
      cursor,
      dataIndex: block.data.length,
      tableColumnIdMap,
      tableWidth,
      tableInlineStyle,
      variables,
    });
  }
}

function compileTableTitleRow(
  titleRow: TableTitleRow,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
  tableWidth: number,
  variables: VariableScope,
): void {
  builder.addCell(context.sheet.id, {
    row: cursor.row,
    column: cursor.column,
    ...compileCellContent(
      interpolateCellValue(titleRow.value, variables),
      createFormulaCompileContext(context.formulaIds ?? new Map(), {
        currentSheetId: context.sheet.id,
        namedRanges: context.namedRangeNames,
      }),
    ),
    style: titleRow.style,
  });

  if (tableWidth > 1) {
    builder.addMerge(context.sheet.id, {
      startRow: cursor.row,
      startColumn: cursor.column,
      endRow: cursor.row,
      endColumn: cursor.column + tableWidth - 1,
    });
  }

  if (titleRow.height !== undefined) {
    builder.setRowHeight(context.sheet.id, {
      row: cursor.row,
      height: titleRow.height,
    });
  }

  cursor.advanceRows();
}

function compileTableDataRows(
  rows: Record<string, unknown>[],
  firstDataIndex: number,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
  leafColumns: TableLeafColumn[],
  tableColumnIdMap: Map<string, number>,
  variables: VariableScope,
  bodyStyle?: StyleValue,
  bandStyle?: StyleValue,
  inlineStyle?: CellStyleDefinition,
  bodyRowHeight?: number,
  hidden?: boolean,
): void {
  for (const [rowOffset, rowData] of rows.entries()) {
    const absoluteRow = cursor.row;
    const dataIndex = firstDataIndex + rowOffset;
    const formulaContext = createTableRowFormulaContext(tableColumnIdMap, absoluteRow, cursor.column, context);

    for (const [columnOffset, column] of leafColumns.entries()) {
      const value = resolveTableCellValue(rowData, column);
      assertTableCellValue(value);
      const compiledContent = compileCellContent(interpolateCellValue(value, variables), formulaContext);
      const staticStyle = column.bodyStyle ?? column.style ?? bandStyle ?? bodyStyle;

      builder.addCell(context.sheet.id, {
        row: absoluteRow,
        column: cursor.column + columnOffset,
        ...compiledContent,
        style: resolveCellStyle(
          staticStyle,
          column.styleResolver?.(compiledContent.value, rowData, dataIndex),
          context,
          `table column "${String(column.id ?? column.title)}"`,
        ),
        inlineStyle,
      });
    }

    if (bodyRowHeight !== undefined) {
      builder.setRowHeight(context.sheet.id, {
        row: absoluteRow,
        height: bodyRowHeight,
      });
    }

    if (hidden === true) {
      builder.setRowHidden(context.sheet.id, {
        row: absoluteRow,
        hidden,
      });
    }

    cursor.advanceRows();
  }
}

function compileTableDataRow(
  rowData: Record<string, unknown>,
  dataIndex: number,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
  leafColumns: TableLeafColumn[],
  tableColumnIdMap: Map<string, number>,
  variables: VariableScope,
  bodyStyle?: StyleValue,
  bandStyle?: StyleValue,
  inlineStyle?: CellStyleDefinition,
  bodyRowHeight?: number,
  hidden?: boolean,
): void {
  compileTableDataRows(
    [rowData],
    dataIndex,
    context,
    cursor,
    builder,
    leafColumns,
    tableColumnIdMap,
    variables,
    bodyStyle,
    bandStyle,
    inlineStyle,
    bodyRowHeight,
    hidden,
  );
}

interface CompileTableSectionRowContext {
  allRows: Record<string, unknown>[];
  allRenderedRows: RenderedTableDataRow[];
  builder: RenderPlanBuilder;
  context: SheetContext;
  currentRenderedRows: RenderedTableDataRow[];
  cursor: LayoutCursor;
  dataIndex: number;
  tableColumnIdMap: Map<string, number>;
  tableWidth: number;
  tableInlineStyle?: CellStyleDefinition;
  variables: VariableScope;
}

function compileTableSectionRow(
  sectionRow: TableSectionRow<Record<string, unknown>>,
  options: CompileTableSectionRowContext,
): void {
  const {
    allRows,
    allRenderedRows,
    builder,
    context,
    currentRenderedRows,
    cursor,
    dataIndex,
    tableColumnIdMap,
    tableWidth,
    tableInlineStyle,
    variables,
  } = options;
  const occupiedColumns = new Set<number>();
  const formulaContext = createTableSectionRowFormulaContext(
    tableColumnIdMap,
    currentRenderedRows,
    allRenderedRows,
    cursor.row,
    cursor.column,
    context,
  );

  if (sectionRow.height !== undefined) {
    builder.setRowHeight(context.sheet.id, {
      row: cursor.row,
      height: sectionRow.height,
    });
  }

  if (sectionRow.hidden !== undefined) {
    builder.setRowHidden(context.sheet.id, {
      row: cursor.row,
      hidden: sectionRow.hidden,
    });
  }

  for (const [cellIndex, cell] of sectionRow.cells.entries()) {
    const columnOffset = resolveSectionCellColumnOffset(cell, cellIndex, tableColumnIdMap, occupiedColumns);
    const colSpan = resolveSectionCellColSpan(cell, columnOffset, tableWidth);
    const value = resolveSectionCellValue(cell, {
      rows: currentRenderedRows.map((row) => row.data),
      allRows,
      dataIndex,
      rowIndex: cursor.row,
    });

    assertTableCellValue(value);
    const compiledContent = compileCellContent(interpolateCellValue(value, variables), formulaContext);

    builder.addCell(context.sheet.id, {
      row: cursor.row,
      column: cursor.column + columnOffset,
      ...compiledContent,
      style: resolveCellStyle(
        cell.style ?? sectionRow.style,
        cell.styleResolver?.(compiledContent.value),
        context,
        `table section cell ${cellIndex + 1}`,
      ),
      inlineStyle: tableInlineStyle,
    });

    if (colSpan > 1) {
      builder.addMerge(context.sheet.id, {
        startRow: cursor.row,
        startColumn: cursor.column + columnOffset,
        endRow: cursor.row,
        endColumn: cursor.column + columnOffset + colSpan - 1,
      });
    }

    for (let offset = columnOffset; offset < columnOffset + colSpan; offset += 1) {
      if (occupiedColumns.has(offset)) {
        throw new ReportEngineError('Table section row cells must not overlap.');
      }

      occupiedColumns.add(offset);
    }
  }

  if (sectionRow.style) {
    for (let offset = 0; offset < tableWidth; offset += 1) {
      if (occupiedColumns.has(offset)) {
        continue;
      }

      builder.addCell(context.sheet.id, {
        row: cursor.row,
        column: cursor.column + offset,
        value: null,
        style: sectionRow.style,
        inlineStyle: tableInlineStyle,
      });
    }
  }

  cursor.advanceRows();
}

function compileTableFooterRow(
  footerRow: TableFooterRow<Record<string, unknown>>,
  options: CompileTableSectionRowContext,
): void {
  compileTableSectionRow(
    {
      type: 'section',
      style: footerRow.style,
      height: footerRow.height,
      cells: footerRow.cells,
    },
    options,
  );
}

function resolveTableFooterRows(
  block: Extract<Block, { type: 'table' }>,
  leafColumns: TableLeafColumn[],
): TableFooterRow<Record<string, unknown>>[] {
  if (block.footerRows) {
    return [...block.footerRows] as TableFooterRow<Record<string, unknown>>[];
  }

  const hasSummary = leafColumns.some((column) => column.summary !== undefined);

  if (!hasSummary) {
    return [];
  }

  return [
    {
      style: block.summaryStyle,
      cells: leafColumns.map((column) => {
        if (column.summary === undefined) {
          return { value: null };
        }

        if (!column.id) {
          throw new ReportEngineError(`Summary column "${column.title}" must include an id.`);
        }

        return {
          columnId: String(column.id),
          style: block.summaryStyle,
          value: createSummaryFormula(String(column.id), column.summary),
        };
      }),
    },
  ];
}

function createSummaryFormula(columnId: string, summary: NonNullable<TableLeafColumn['summary']>): CellContent {
  if (summary && typeof summary === 'object') {
    return summary;
  }

  const range = {
    type: 'range' as const,
    startId: columnId,
    endId: columnId,
    scope: 'allRows' as const,
  };

  switch (summary) {
    case 'sum':
      return { type: 'sum', range };
    case 'count':
      return { type: 'call', name: 'COUNT', args: [range] };
    case 'average':
      return { type: 'call', name: 'AVERAGE', args: [range] };
    default:
      return assertNeverSummary(summary);
  }
}

function resolveCellStyle(
  staticStyle: StyleValue | undefined,
  dynamicStyle: StyleValue | undefined,
  context: SheetContext,
  label: string,
): StyleValue | undefined {
  if (dynamicStyle === undefined) {
    return staticStyle;
  }

  if (typeof dynamicStyle === 'string') {
    if (!context.styles || !Object.prototype.hasOwnProperty.call(context.styles, dynamicStyle)) {
      throw new ReportEngineError(`${label} styleResolver returned unknown style "${dynamicStyle}".`);
    }

    return dynamicStyle;
  }

  if (typeof dynamicStyle !== 'object' || dynamicStyle === null || Array.isArray(dynamicStyle)) {
    throw new ReportEngineError(`${label} styleResolver must return a style value.`);
  }

  return dynamicStyle;
}

function createTableInlineStyle(border: TableBorderDefinition | undefined): CellStyleDefinition | undefined {
  if (!border) {
    return undefined;
  }

  return {
    border: typeof border === 'string' ? createAllSidesBorder(border) : border,
  };
}

function createAllSidesBorder(style: ExcelJS.BorderStyle): Partial<ExcelJS.Borders> {
  return {
    top: { style },
    right: { style },
    bottom: { style },
    left: { style },
  };
}

interface RenderedTableDataRow {
  data: Record<string, unknown>;
  row: number;
}

function resolveSectionCellColumnOffset(
  cell: TableSectionCell<Record<string, unknown>>,
  cellIndex: number,
  tableColumnIdMap: Map<string, number>,
  occupiedColumns: Set<number>,
): number {
  if (cell.column !== undefined) {
    return cell.column - 1;
  }

  if (cell.columnId !== undefined) {
    const offset = tableColumnIdMap.get(cell.columnId);

    if (offset === undefined) {
      throw new ReportEngineError(`Table section row references unknown columnId "${cell.columnId}".`);
    }

    return offset;
  }

  let offset = cellIndex === 0 ? 0 : Math.max(...occupiedColumns) + 1;

  while (occupiedColumns.has(offset)) {
    offset += 1;
  }

  return offset;
}

function resolveSectionCellColSpan(
  cell: TableSectionCell<Record<string, unknown>>,
  columnOffset: number,
  tableWidth: number,
): number {
  const colSpan = cell.colSpan === 'remaining' ? tableWidth - columnOffset : (cell.colSpan ?? 1);

  if (columnOffset < 0 || columnOffset >= tableWidth || columnOffset + colSpan > tableWidth) {
    throw new ReportEngineError('Table section row cell range exceeds table width.');
  }

  return colSpan;
}

function resolveSectionCellValue(
  cell: TableSectionCell<Record<string, unknown>>,
  context: TableSectionCellContext<Record<string, unknown>>,
): unknown {
  if (typeof cell.value === 'function') {
    return cell.value(context);
  }

  return cell.value ?? null;
}

function collectTableDataRows(block: Extract<Block, { type: 'table' }>): Record<string, unknown>[] {
  if (!Array.isArray(block.data)) {
    return [];
  }

  return block.data.filter((item) => !isTableSectionRow(item)) as Record<string, unknown>[];
}

function isTableSectionRow(item: unknown): item is TableSectionRow<Record<string, unknown>> {
  return typeof item === 'object' && item !== null && 'type' in item && item.type === 'section';
}

function createTableColumnIdMap(columns: TableLeafColumn[]): Map<string, number> {
  const idMap = new Map<string, number>();

  for (const [columnOffset, column] of columns.entries()) {
    if (!column.id) {
      continue;
    }

    if (idMap.has(String(column.id))) {
      throw new ReportEngineError(`Duplicate formula cell id "${String(column.id)}".`);
    }

    idMap.set(String(column.id), columnOffset);
  }

  return idMap;
}

function createTableRowFormulaContext(
  columnIdMap: Map<string, number>,
  row: number,
  firstColumn: number,
  context: SheetContext,
): FormulaCompileContext {
  return {
    resolveCellId(id: string, sheetId?: string): string {
      if (sheetId && sheetId !== context.sheet.id) {
        const address = context.formulaIds?.get(createFormulaId(sheetId, id));

        if (!address) {
          throw new ReportEngineError(`Formula references unknown cell id "${id}".`);
        }

        return formatCellReference(address, context.sheet.id);
      }

      const columnOffset = columnIdMap.get(id);

      if (columnOffset === undefined) {
        throw new ReportEngineError(`Formula references unknown cell id "${id}".`);
      }

      return formatCellAddress({
        row,
        column: firstColumn + columnOffset,
      });
    },
    resolveRangeIds(startId: string, endId: string, sheetId?: string, scope?: FormulaRangeScope): string {
      if (scope) {
        throw new ReportEngineError('Scoped formula ranges are only supported inside table section rows.');
      }

      if (sheetId && sheetId !== context.sheet.id) {
        const start = context.formulaIds?.get(createFormulaId(sheetId, startId));
        const end = context.formulaIds?.get(createFormulaId(sheetId, endId));

        if (!start) {
          throw new ReportEngineError(`Formula references unknown range start id "${startId}".`);
        }

        if (!end) {
          throw new ReportEngineError(`Formula references unknown range end id "${endId}".`);
        }

        if (end.row < start.row || end.column < start.column) {
          throw new ReportEngineError('Formula range end id must resolve after start id.');
        }

        return [formatCellReference(start, context.sheet.id), formatCellReference(end, context.sheet.id)].join(':');
      }

      const startColumnOffset = columnIdMap.get(startId);
      const endColumnOffset = columnIdMap.get(endId);

      if (startColumnOffset === undefined) {
        throw new ReportEngineError(`Formula references unknown range start id "${startId}".`);
      }

      if (endColumnOffset === undefined) {
        throw new ReportEngineError(`Formula references unknown range end id "${endId}".`);
      }

      if (endColumnOffset < startColumnOffset) {
        throw new ReportEngineError('Formula range end id must resolve after start id.');
      }

      return [
        formatCellAddress({ row, column: firstColumn + startColumnOffset }),
        formatCellAddress({ row, column: firstColumn + endColumnOffset }),
      ].join(':');
    },
    resolveNamedRange(name: string): string {
      if (context.namedRangeNames && !context.namedRangeNames.has(name)) {
        throw new ReportEngineError(`Formula references unknown named range "${name}".`);
      }

      return name;
    },
  };
}

function createTableSectionRowFormulaContext(
  columnIdMap: Map<string, number>,
  currentRows: RenderedTableDataRow[],
  allRows: RenderedTableDataRow[],
  row: number,
  firstColumn: number,
  context: SheetContext,
): FormulaCompileContext {
  const rowContext = createTableRowFormulaContext(columnIdMap, row, firstColumn, context);

  return {
    resolveCellId: rowContext.resolveCellId,
    resolveNamedRange: rowContext.resolveNamedRange,
    resolveRangeIds(startId: string, endId: string, sheetId?: string, scope?: FormulaRangeScope): string {
      if (!scope) {
        return rowContext.resolveRangeIds(startId, endId, sheetId);
      }

      if (sheetId && sheetId !== context.sheet.id) {
        throw new ReportEngineError('Scoped formula ranges must reference the current table sheet.');
      }

      const rows = scope === 'allRows' ? allRows : currentRows;
      const startColumnOffset = columnIdMap.get(startId);
      const endColumnOffset = columnIdMap.get(endId);

      if (startColumnOffset === undefined) {
        throw new ReportEngineError(`Formula references unknown range start id "${startId}".`);
      }

      if (endColumnOffset === undefined) {
        throw new ReportEngineError(`Formula references unknown range end id "${endId}".`);
      }

      if (endColumnOffset < startColumnOffset) {
        throw new ReportEngineError('Formula range end id must resolve after start id.');
      }

      if (rows.length === 0) {
        return '0';
      }

      return formatTableRowRanges(rows, firstColumn + startColumnOffset, firstColumn + endColumnOffset);
    },
  };
}

function formatTableRowRanges(rows: RenderedTableDataRow[], startColumn: number, endColumn: number): string {
  const sortedRows = [...rows].sort((left, right) => left.row - right.row);
  const ranges: string[] = [];
  let startRow = sortedRows[0]?.row;
  let endRow = startRow;

  for (const row of sortedRows.slice(1)) {
    if (endRow !== undefined && row.row === endRow + 1) {
      endRow = row.row;
      continue;
    }

    if (startRow !== undefined && endRow !== undefined) {
      ranges.push(formatTableRowRange(startRow, endRow, startColumn, endColumn));
    }

    startRow = row.row;
    endRow = row.row;
  }

  if (startRow !== undefined && endRow !== undefined) {
    ranges.push(formatTableRowRange(startRow, endRow, startColumn, endColumn));
  }

  return ranges.join(',');
}

function formatTableRowRange(startRow: number, endRow: number, startColumn: number, endColumn: number): string {
  return [
    formatCellAddress({ row: startRow, column: startColumn }),
    formatCellAddress({ row: endRow, column: endColumn }),
  ].join(':');
}

function registerCellId(idMap: Map<string, CellAddress>, registryKey: string, address: CellAddress): void {
  if (idMap.has(registryKey)) {
    throw new ReportEngineError(`Duplicate formula cell id "${registryKey}".`);
  }

  idMap.set(registryKey, address);
}

function createBlockVariableScope(context: SheetContext, block: Block): VariableScope {
  return {
    workbook: context.variables.workbook,
    sheet: context.variables.sheet,
    block: block.context,
  };
}

function resolveTableCellValue(row: Record<string, unknown>, column: TableLeafColumn): unknown {
  if (column.accessor) {
    return column.accessor(row);
  }

  const value = column.id ? row[String(column.id)] : null;
  return value === undefined ? null : value;
}

function assertTableCellValue(value: unknown): asserts value is CellContent {
  if (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean' ||
    value instanceof Date ||
    isFormulaDefinition(value)
  ) {
    return;
  }

  throw new ReportEngineError('Table cell values must resolve to a supported cell value.');
}

interface HeaderMatrixCell {
  title: string;
  style?: StyleValue;
  rowOffset: number;
  columnOffset: number;
  rowSpan: number;
  colSpan: number;
}

type TableColumnNode = Extract<Block, { type: 'table' }>['columns'][number];
type TableLeafColumn = TableColumnNode & { children?: undefined };

function calculateHeaderDepth(columns: readonly TableColumnNode[]): number {
  return Math.max(
    ...columns.map((column) => {
      if (column.children && column.children.length > 0) {
        return getChildrenRowOffset(column) + calculateHeaderDepth(column.children);
      }

      return 1;
    }),
  );
}

function flattenLeafColumns(columns: readonly TableColumnNode[]): TableLeafColumn[] {
  return columns.flatMap((column) => {
    if (column.children && column.children.length > 0) {
      return flattenLeafColumns(column.children);
    }

    return [column as TableLeafColumn];
  });
}

function buildHeaderMatrix(columns: readonly TableColumnNode[], headerDepth: number): HeaderMatrixCell[] {
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
  const childrenRowOffset = isParent ? getChildrenRowOffset(column) : 0;
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

  if (!isParent) {
    return 1;
  }

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

function getChildrenRowOffset(column: TableColumnNode): number {
  return column.childrenRowOffset ?? 1;
}

function countLeafColumns(columns: readonly TableColumnNode[]): number {
  return columns.reduce((total, column) => {
    if (column.children && column.children.length > 0) {
      return total + countLeafColumns(column.children);
    }

    return total + 1;
  }, 0);
}

function assertNeverSummary(value: never): never {
  throw new ReportEngineError(`Unsupported table summary "${String(value)}".`);
}
