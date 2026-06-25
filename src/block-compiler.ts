import type { CompileContext } from './compile-context';
import { CompileError, ReportEngineError } from './errors';
import {
  createGridFormulaContext,
  createTableRowFormulaContext,
  createTableSectionFormulaContext,
  type RenderedDataRow,
} from './formula-context';
import { compileCellContent, isFormulaDefinition } from './formula-engine';
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
  resolveSectionCellColSpan,
  resolveSectionCellColumnOffset,
  resolveSectionCellValue,
} from './helpers/utils';
import type { SheetWriter } from './sheet-writer';
import type {
  Block,
  CellContent,
  CellStyleDefinition,
  GridCell,
  StyleValue,
  TableFooterRow,
  TableGroup,
  TableLeafColumn,
  TableSectionRow,
} from './types';
import { interpolateCellValue, interpolateVariables, type VariableScope } from './variable-engine';

// ─── Public entry point ──────────────────────────────────────────────────────

/**
 * Compile một block vào `writer`, cập nhật `row` (cursor dòng hiện tại)
 * và trả về số dòng đã tiêu thụ.
 *
 * Thay vì nhận mutable LayoutCursor, hàm nhận `row` và trả về row mới.
 * Điều này làm rõ side-effect: caller biết chính xác bao nhiêu dòng bị dùng.
 */
export function compileBlock(block: Block, context: CompileContext, writer: SheetWriter, startRow: number): number {
  switch (block.type) {
    case 'title':
    case 'text':
      return compileTextBlock(block, context, writer, startRow);
    case 'spacer':
      return startRow + (block.rows ?? 1);
    case 'divider':
      return compileDividerBlock(block, writer, startRow);
    case 'grid':
      return compileGridBlock(block, context, writer, startRow);
    case 'table':
    case 'table-groups':
      return compileTableBlock(block, context, writer, startRow);
    default:
      throw new ReportEngineError(`Unknown block type "${(block as Block).type}" in sheet "${context.sheet.id}".`);
  }
}

// ─── Text / Title ─────────────────────────────────────────────────────────────

function compileTextBlock(
  block: Extract<Block, { type: 'title' | 'text' }>,
  context: CompileContext,
  writer: SheetWriter,
  row: number,
): number {
  const variables = blockVariables(context, block);
  const colSpan = resolveColSpan(block.colSpan, 0, context.sheetColumnCount);

  writer.addCell({ row, column: 1, value: interpolateVariables(block.text, variables), style: block.style });

  if (colSpan > 1) {
    writer.addMerge({ startRow: row, startColumn: 1, endRow: row, endColumn: colSpan });
  }

  if (block.height !== undefined) {
    writer.setRowHeight({ row, height: block.height });
  }

  return row + 1;
}

// ─── Divider ──────────────────────────────────────────────────────────────────

function compileDividerBlock(
  block: Extract<Block, { type: 'divider' }>,
  writer: SheetWriter,
  startRow: number,
): number {
  const rows = block.rows ?? 1;

  for (let offset = 0; offset < rows; offset++) {
    writer.addCell({ row: startRow + offset, column: 1, value: '', style: block.style });
  }

  return startRow + rows;
}

// ─── Grid ─────────────────────────────────────────────────────────────────────

interface GridPlacement {
  cell: GridCell;
  row: number;
  column: number;
  rowSpan: number;
  colSpan: number;
  rowOffset: number;
}

function compileGridBlock(
  block: Extract<Block, { type: 'grid' }>,
  context: CompileContext,
  writer: SheetWriter,
  startRow: number,
): number {
  const occupied = new Set<string>();
  const placements: GridPlacement[] = [];
  const variables = blockVariables(context, block);
  let rowExtent = block.rows.length;

  // Pass 1: layout — tính vị trí và đăng ký IDs ngay vào registry
  for (const [rowOffset, gridRow] of block.rows.entries()) {
    const absRow = startRow + rowOffset;

    if (gridRow.height !== undefined) {
      writer.setRowHeight({ row: absRow, height: gridRow.height });
    }

    let colOffset = 0;

    for (const cell of gridRow.cells) {
      while (occupied.has(gridOccupancyKey(rowOffset, colOffset))) colOffset++;

      const colSpan = resolveColSpan(cell.colSpan, colOffset, context.sheetColumnCount);
      const rowSpan = cell.rowSpan ?? 1;
      assertGridCellDoesNotOverlap(occupied, rowOffset, colOffset, rowSpan, colSpan);

      const absCol = 1 + colOffset;

      // Đăng ký ngay vào registry — single-pass, không cần pre-pass
      if (cell.id) {
        context.registry.register(context.sheet.id, context.sheet.name, cell.id, absRow, absCol);
      }

      placements.push({ cell, row: absRow, column: absCol, rowSpan, colSpan, rowOffset });
      markGridOccupied(occupied, rowOffset, colOffset, rowSpan, colSpan);
      rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
      colOffset += colSpan;
    }
  }

  // Pass 2: render — sau khi tất cả IDs đã đăng ký, compile formula
  const formulaCtx = createGridFormulaContext(context.registry, context.sheet.id);

  for (const p of placements) {
    const compiled = compileCellContent(interpolateCellValue(p.cell.value, variables), formulaCtx);

    writer.addCell({
      row: p.row,
      column: p.column,
      ...compiled,
      formulaResult: isFormulaDefinition(p.cell.value) ? p.cell.formulaResult : undefined,
      style: resolveStyle(
        p.cell.style,
        p.cell.styleResolver?.(compiled.value),
        context,
        `grid cell "${p.cell.id ?? `${p.row}:${p.column}`}"`,
      ),
    });

    if (p.cell.width !== undefined) {
      writer.setColumnWidth({ column: p.column, width: p.cell.width });
    }

    if (p.rowSpan > 1 || p.colSpan > 1) {
      writer.addMerge({
        startRow: p.row,
        startColumn: p.column,
        endRow: p.row + p.rowSpan - 1,
        endColumn: p.column + p.colSpan - 1,
      });
    }
  }

  return startRow + rowExtent;
}

// ─── Table ────────────────────────────────────────────────────────────────────

function compileTableBlock(
  block: Extract<Block, { type: 'table' | 'table-groups' }>,
  context: CompileContext,
  writer: SheetWriter,
  startRow: number,
): number {
  const variables = blockVariables(context, block);
  const headerDepth = calculateTableHeaderDepth(block.columns);
  const leafColumns = flattenColumns(block.columns);
  const columnIdMap = createTableColumnIdMap(leafColumns);
  const tableWidth = leafColumns.length;
  const allDataRows = collectTableDataRows(block);
  const inlineStyle = createTableInlineStyle(block.border);
  const footerRows = resolveFooterRows(block, leafColumns);

  let row = startRow;

  // Header row heights
  for (const [offset, height] of (block.headerRowHeights ?? []).entries()) {
    writer.setRowHeight({ row: row + offset, height });
  }

  // Header cells
  for (const hCell of buildHeaderMatrix(block.columns, headerDepth)) {
    writer.addCell({
      row: row + hCell.rowOffset,
      column: 1 + hCell.columnOffset,
      value: interpolateVariables(hCell.title, variables),
      style: hCell.style ?? block.headerStyle,
      inlineStyle,
    });

    if (hCell.rowSpan > 1 || hCell.colSpan > 1) {
      writer.addMerge({
        startRow: row + hCell.rowOffset,
        startColumn: 1 + hCell.columnOffset,
        endRow: row + hCell.rowOffset + hCell.rowSpan - 1,
        endColumn: 1 + hCell.columnOffset + hCell.colSpan - 1,
      });
    }
  }

  // Column widths / hidden
  for (const [offset, col] of leafColumns.entries()) {
    if (col.width !== undefined) writer.setColumnWidth({ column: 1 + offset, width: col.width });
    if (col.hidden !== undefined) writer.setColumnHidden({ column: 1 + offset, hidden: col.hidden });
  }

  row += headerDepth;

  const allRendered: RenderedDataRow[] = [];

  if (block.type === 'table') {
    row = writeDataRows(block.data as Record<string, unknown>[], 0, allRendered, undefined, row, {
      block,
      context,
      writer,
      leafColumns,
      columnIdMap,
      inlineStyle,
      variables,
    });
  } else {
    let dataIndex = 0;

    for (const group of block.groups as TableGroup[]) {
      const groupRendered: RenderedDataRow[] = [];

      for (const sectionRow of group.headerRows ?? []) {
        row = writeSectionRow(sectionRow, {
          allDataRows,
          allRendered,
          currentRendered: groupRendered,
          dataIndex,
          row,
          tableWidth,
          columnIdMap,
          inlineStyle,
          variables,
          context,
          writer,
        });
      }

      row = writeDataRows(group.data as Record<string, unknown>[], dataIndex, allRendered, groupRendered, row, {
        block,
        context,
        writer,
        leafColumns,
        columnIdMap,
        inlineStyle,
        variables,
      });

      for (const sectionRow of group.footerRows ?? []) {
        row = writeSectionRow(sectionRow, {
          allDataRows,
          allRendered,
          currentRendered: groupRendered,
          dataIndex: dataIndex + group.data.length,
          row,
          tableWidth,
          columnIdMap,
          inlineStyle,
          variables,
          context,
          writer,
        });
      }

      dataIndex += group.data.length;
    }
  }

  // Table-level footer rows
  for (const footerRow of footerRows) {
    row = writeSectionRow(
      { style: footerRow.style, height: footerRow.height, cells: footerRow.cells },
      {
        allDataRows,
        allRendered,
        currentRendered: allRendered,
        dataIndex: allDataRows.length,
        row,
        tableWidth,
        columnIdMap,
        inlineStyle,
        variables,
        context,
        writer,
      },
    );
  }

  return row;
}

// ─── Table helpers ────────────────────────────────────────────────────────────

interface DataRowsOptions {
  block: Extract<Block, { type: 'table' | 'table-groups' }>;
  context: CompileContext;
  writer: SheetWriter;
  leafColumns: TableLeafColumn[];
  columnIdMap: Map<string, number>;
  inlineStyle?: CellStyleDefinition;
  variables: VariableScope;
}

function writeDataRows(
  rows: Record<string, unknown>[],
  firstDataIndex: number,
  allRendered: RenderedDataRow[],
  groupRendered: RenderedDataRow[] | undefined,
  startRow: number,
  opts: DataRowsOptions,
): number {
  let row = startRow;

  for (const [offset, rowData] of rows.entries()) {
    const dataIndex = firstDataIndex + offset;
    const formulaCtx = createTableRowFormulaContext(
      opts.columnIdMap,
      row,
      1,
      opts.context.registry,
      opts.context.sheet.id,
    );
    const bandStyle = dataIndex % 2 === 0 ? opts.block.oddRowStyle : opts.block.evenRowStyle;

    for (const [colOffset, col] of opts.leafColumns.entries()) {
      const rawValue = resolveTableCellValue(rowData, col);
      assertTableCellValue(rawValue);
      const compiled = compileCellContent(interpolateCellValue(rawValue, opts.variables), formulaCtx);

      opts.writer.addCell({
        row,
        column: 1 + colOffset,
        ...compiled,
        style: resolveStyle(
          col.bodyStyle ?? col.style ?? bandStyle ?? opts.block.bodyStyle,
          col.styleResolver?.(compiled.value, rowData, dataIndex),
          opts.context,
          `table column "${String(col.id ?? col.title)}"`,
        ),
        inlineStyle: opts.inlineStyle,
      });
    }

    if (opts.block.bodyRowHeight !== undefined) {
      opts.writer.setRowHeight({ row, height: opts.block.bodyRowHeight });
    }

    if (opts.block.rowHidden?.(rowData, dataIndex) === true) {
      opts.writer.setRowHidden({ row, hidden: true });
    }

    const rendered: RenderedDataRow = { data: rowData, row };
    groupRendered?.push(rendered);
    allRendered.push(rendered);
    row++;
  }

  return row;
}

interface SectionRowOptions {
  allDataRows: Record<string, unknown>[];
  allRendered: RenderedDataRow[];
  currentRendered: RenderedDataRow[];
  dataIndex: number;
  row: number;
  tableWidth: number;
  columnIdMap: Map<string, number>;
  inlineStyle?: CellStyleDefinition;
  variables: VariableScope;
  context: CompileContext;
  writer: SheetWriter;
}

interface SectionCellPlacement {
  cell: TableSectionRow['cells'][number];
  cellIndex: number;
  columnOffset: number;
  colSpan: number;
}

function writeSectionRow(sectionRow: TableSectionRow, opts: SectionRowOptions): number {
  const { row, context, writer } = opts;
  const occupied = new Set<number>();
  const placements: SectionCellPlacement[] = [];

  if (sectionRow.height !== undefined) writer.setRowHeight({ row, height: sectionRow.height });
  if (sectionRow.hidden !== undefined) writer.setRowHidden({ row, hidden: sectionRow.hidden });

  for (const [cellIndex, cell] of sectionRow.cells.entries()) {
    const colOffset = resolveSectionCellColumnOffset(cell, cellIndex, opts.columnIdMap, occupied);
    const colSpan = resolveSectionCellColSpan(cell, colOffset, opts.tableWidth);

    if (cell.id) {
      context.registry.register(context.sheet.id, context.sheet.name, cell.id, row, 1 + colOffset);
    }

    for (let o = colOffset; o < colOffset + colSpan; o++) {
      if (occupied.has(o)) throw new CompileError('Table section row cells must not overlap.');
      occupied.add(o);
    }

    placements.push({ cell, cellIndex, columnOffset: colOffset, colSpan });
  }

  const formulaCtx = createTableSectionFormulaContext(
    opts.columnIdMap,
    opts.currentRendered,
    opts.allRendered,
    row,
    1,
    context.registry,
    context.sheet.id,
  );

  for (const { cell, cellIndex, columnOffset, colSpan } of placements) {
    const rawValue = resolveSectionCellValue(cell, {
      rows: opts.currentRendered.map((r) => r.data),
      allRows: opts.allDataRows,
      dataIndex: opts.dataIndex,
      rowIndex: row,
    });

    assertTableCellValue(rawValue);
    const compiled = compileCellContent(interpolateCellValue(rawValue, opts.variables), formulaCtx);

    writer.addCell({
      row,
      column: 1 + columnOffset,
      ...compiled,
      style: resolveStyle(
        cell.style ?? sectionRow.style,
        cell.styleResolver?.(compiled.value),
        context,
        `table section cell ${cellIndex + 1}`,
      ),
      inlineStyle: opts.inlineStyle,
    });

    if (colSpan > 1) {
      writer.addMerge({
        startRow: row,
        startColumn: 1 + columnOffset,
        endRow: row,
        endColumn: 1 + columnOffset + colSpan - 1,
      });
    }
  }

  // Fill remaining cells với sectionRow.style nếu có
  if (sectionRow.style) {
    for (let o = 0; o < opts.tableWidth; o++) {
      if (!occupied.has(o)) {
        writer.addCell({ row, column: 1 + o, value: null, style: sectionRow.style, inlineStyle: opts.inlineStyle });
      }
    }
  }

  return row + 1;
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function resolveFooterRows(
  block: Extract<Block, { type: 'table' | 'table-groups' }>,
  leafColumns: TableLeafColumn[],
): TableFooterRow[] {
  if (block.footerRows) return [...block.footerRows] as TableFooterRow[];

  const hasSummary = leafColumns.some((col) => col.summary !== undefined);
  if (!hasSummary) return [];

  return [
    {
      style: block.summaryStyle,
      cells: leafColumns.map((col) => {
        if (col.summary === undefined) return { value: null };

        if (!col.id) {
          throw new ReportEngineError(`Summary column "${col.title}" must include an id.`);
        }

        return {
          columnId: String(col.id),
          style: block.summaryStyle,
          value: createSummaryFormula(String(col.id), col.summary),
        };
      }),
    },
  ];
}

function resolveStyle(
  staticStyle: StyleValue | undefined,
  dynamicStyle: StyleValue | undefined,
  context: CompileContext,
  label: string,
): StyleValue | undefined {
  if (dynamicStyle === undefined) return staticStyle;

  if (typeof dynamicStyle === 'string') {
    const styles = context.workbook.styles;

    if (!styles || !Object.prototype.hasOwnProperty.call(styles, dynamicStyle)) {
      throw new ReportEngineError(`${label} styleResolver returned unknown style "${dynamicStyle}".`);
    }

    return dynamicStyle;
  }

  if (typeof dynamicStyle !== 'object' || dynamicStyle === null || Array.isArray(dynamicStyle)) {
    throw new ReportEngineError(`${label} styleResolver must return a style value.`);
  }

  return dynamicStyle;
}

function blockVariables(context: CompileContext, block: Block): VariableScope {
  return {
    workbook: context.variables.workbook,
    sheet: context.variables.sheet,
    block: block.context,
  };
}

function resolveTableCellValue(row: Record<string, unknown>, col: TableLeafColumn): unknown {
  if (col.accessor) return col.accessor(row);
  const value = col.id ? row[String(col.id)] : null;
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
