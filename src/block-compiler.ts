import { ReportEngineError } from './errors';
import type {
  Block,
  CellContent,
  CellStyleDefinition,
  FormulaRangeScope,
  GridCell,
  SheetDefinition,
  StyleRegistry,
  StyleValue,
  TableFooterRow,
  TableGroup,
  TableLeafColumn,
  TableSectionRow,
  WorkbookDefinition,
} from './types';
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
import {
  assertGridCellDoesNotOverlap,
  buildHeaderMatrix,
  calculateTableHeaderDepth,
  collectTableDataRows,
  createSummaryFormula,
  createTableColumnIdMap,
  createTableInlineStyle,
  flattenColumns,
  gridOccupancyKey,
  markGridOccupied,
  resolveColSpan,
  resolveSectionCellColumnOffset,
  resolveSectionCellColSpan,
  resolveSectionCellValue,
} from './helpers/utils';
import { interpolateCellValue, interpolateVariables, type VariableScope } from './variable-engine';

export interface SheetContext {
  workbook: WorkbookDefinition;
  sheet: SheetDefinition;
  styles?: StyleRegistry;
  variables: VariableScope;
  formulaIds?: Map<string, CellAddress>;
  sheetColumnCount: number;
}

function compileTextLikeBlock(
  block: Extract<Block, { type: 'title' | 'text' }>,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
): void {
  const variables = createBlockVariableScope(context, block);
  const colSpan = resolveColSpan(block.colSpan, 0, context.sheetColumnCount);

  builder.addCell(context.sheet.id, {
    row: cursor.row,
    column: cursor.column,
    value: interpolateVariables(block.text, variables),
    style: block.style,
  });

  if (colSpan > 1) {
    builder.addMerge(context.sheet.id, {
      startRow: cursor.row,
      startColumn: cursor.column,
      endRow: cursor.row,
      endColumn: cursor.column + colSpan - 1,
    });
  }

  if (block.height !== undefined) {
    builder.setRowHeight(context.sheet.id, {
      row: cursor.row,
      height: block.height,
    });
  }

  cursor.advanceRows();
}

export function compileBlock(
  block: Block,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
): void {
  switch (block.type) {
    case 'title':
    case 'text':
      compileTextLikeBlock(block, context, cursor, builder);
      return;
    case 'spacer':
      cursor.advanceRows(block.rows ?? 1);
      return;
    case 'divider':
      compileDividerBlock(block, context, cursor, builder);
      return;
    case 'grid':
      compileGridBlock(block, context, cursor, builder);
      return;
    case 'table':
    case 'table-groups':
      compileTableBlock(block, context, cursor, builder);
      return;
    default:
      throw new ReportEngineError(`Unknown block type "${(block as Block).type}" in sheet "${context.sheet.id}".`);
  }
}

function compileDividerBlock(
  block: Extract<Block, { type: 'divider' }>,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
): void {
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
      while (occupied.has(gridOccupancyKey(rowOffset, columnOffset))) {
        columnOffset += 1;
      }

      const colSpan = resolveColSpan(cell.colSpan, columnOffset, context.sheetColumnCount);
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

      markGridOccupied(occupied, rowOffset, columnOffset, rowSpan, colSpan);
      rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
      columnOffset += colSpan;
    }
  }

  const formulaContext = createFormulaCompileContext(
    context.formulaIds ?? createGridCellIdMap(placements, context.sheet),
    {
      currentSheetId: context.sheet.id,
    },
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

    const registryKey = createFormulaId(sheet.id, placement.cell.id);

    if (idMap.has(registryKey)) {
      throw new ReportEngineError(`Duplicate formula cell id "${registryKey}".`);
    }

    idMap.set(registryKey, {
      row: placement.row,
      column: placement.column,
      sheetId: sheet.id,
      sheetName: sheet.name,
    });
  }

  return idMap;
}

function compileTableBlock(
  block: Extract<Block, { type: 'table' | 'table-groups' }>,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
): void {
  const variables = createBlockVariableScope(context, block);
  const headerDepth = calculateTableHeaderDepth(block.columns);
  const headerCells = buildHeaderMatrix(block.columns, headerDepth);
  const leafColumns = flattenColumns(block.columns);
  const tableColumnIdMap = createTableColumnIdMap(leafColumns);
  const tableWidth = leafColumns.length;
  const allRows = collectTableDataRows(block);
  const tableInlineStyle = createTableInlineStyle(block.border);
  const footerRows = resolveTableFooterRows(block, leafColumns);

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

  const allRenderedRows: RenderedTableDataRow[] = [];

  if (block.type === 'table') {
    compileTableDataRowsWithTracking(block.data as readonly Record<string, unknown>[], 0, {
      allRenderedRows,
      block,
      builder,
      context,
      cursor,
      leafColumns,
      tableColumnIdMap,
      tableInlineStyle,
      variables,
    });
  } else {
    let dataIndex = 0;

    for (const group of block.groups as readonly TableGroup[]) {
      const groupRenderedRows: RenderedTableDataRow[] = [];

      for (const sectionRow of group.headerRows ?? []) {
        compileTableSectionRow(sectionRow, {
          allRows,
          allRenderedRows,
          builder,
          context,
          currentRenderedRows: groupRenderedRows,
          cursor,
          dataIndex,
          tableColumnIdMap,
          tableWidth,
          tableInlineStyle,
          variables,
        });
      }

      compileTableDataRowsWithTracking(group.data, dataIndex, {
        allRenderedRows,
        block,
        builder,
        context,
        cursor,
        leafColumns,
        tableColumnIdMap,
        tableInlineStyle,
        trackedRows: groupRenderedRows,
        variables,
      });

      for (const sectionRow of group.footerRows ?? []) {
        compileTableSectionRow(sectionRow, {
          allRows,
          allRenderedRows,
          builder,
          context,
          currentRenderedRows: groupRenderedRows,
          cursor,
          dataIndex: dataIndex + group.data.length,
          tableColumnIdMap,
          tableWidth,
          tableInlineStyle,
          variables,
        });
      }

      dataIndex += group.data.length;
    }
  }

  for (const footerRow of footerRows) {
    compileTableSectionRow(
      {
        style: footerRow.style,
        height: footerRow.height,
        cells: footerRow.cells,
      },
      {
        allRows,
        allRenderedRows,
        builder,
        context,
        currentRenderedRows: allRenderedRows,
        cursor,
        dataIndex: allRows.length,
        tableColumnIdMap,
        tableWidth,
        tableInlineStyle,
        variables,
      },
    );
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
  const absoluteRow = cursor.row;
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

function compileTableSectionRow(sectionRow: TableSectionRow, options: CompileTableSectionRowContext): void {
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

function resolveTableFooterRows(
  block: Extract<Block, { type: 'table' | 'table-groups' }>,
  leafColumns: TableLeafColumn[],
): TableFooterRow[] {
  if (block.footerRows) {
    return [...block.footerRows] as TableFooterRow[];
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

interface RenderedTableDataRow {
  data: Record<string, unknown>;
  row: number;
}

interface CompileTableDataRowsWithTrackingContext {
  allRenderedRows: RenderedTableDataRow[];
  block: Extract<Block, { type: 'table' | 'table-groups' }>;
  builder: RenderPlanBuilder;
  context: SheetContext;
  cursor: LayoutCursor;
  leafColumns: TableLeafColumn[];
  tableColumnIdMap: Map<string, number>;
  tableInlineStyle?: CellStyleDefinition;
  trackedRows?: RenderedTableDataRow[];
  variables: VariableScope;
}

function compileTableDataRowsWithTracking(
  rows: readonly Record<string, unknown>[],
  firstDataIndex: number,
  options: CompileTableDataRowsWithTrackingContext,
): void {
  for (const [rowOffset, rowData] of rows.entries()) {
    const dataIndex = firstDataIndex + rowOffset;
    const rowIndex = options.cursor.row;

    compileTableDataRow(
      rowData,
      dataIndex,
      options.context,
      options.cursor,
      options.builder,
      options.leafColumns,
      options.tableColumnIdMap,
      options.variables,
      options.block.bodyStyle,
      dataIndex % 2 === 0 ? options.block.oddRowStyle : options.block.evenRowStyle,
      options.tableInlineStyle,
      options.block.bodyRowHeight,
      options.block.rowHidden?.(rowData, dataIndex),
    );

    const renderedRow = {
      data: rowData,
      row: rowIndex,
    };

    options.trackedRows?.push(renderedRow);
    options.allRenderedRows.push(renderedRow);
  }
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
