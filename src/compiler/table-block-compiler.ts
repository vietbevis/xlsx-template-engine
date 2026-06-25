import { CompileError } from '../errors';
import { FormulaResolver, type RenderedDataRow } from '../formula/formula-resolver';
import { FormulaCompiler } from '../formula/formula-compiler';
import { writeCell, writeMerge } from '../utils/exceljs-writer';
import { resolveStyle } from '../styles/style-resolver';
import {
  assertTableCellValue,
  buildHeaderMatrix,
  calculateTableHeaderDepth,
  createTableColumnIdMap,
  createTableInlineStyle,
  flattenColumns,
  resolveFooterRows,
  resolveSectionCellColSpan,
  resolveSectionCellColumnOffset,
  resolveSectionCellValue,
  resolveTableCellValue,
} from '../utils/table-utils';
import type { Block, TableGroup, TableSectionCell, TableSectionRow } from '../types';
import type { CompileContext } from './block-dispatcher';

export function compileTableBlock(
  block: Extract<Block, { type: 'table' }>,
  context: CompileContext,
  startRow: number,
): number {
  return new TableBlockCompiler(block, context).compile(startRow);
}

interface SectionCellPlacement {
  cell: TableSectionCell;
  cellIndex: number;
  columnOffset: number;
  colSpan: number;
}

const COL_ORIGIN = 1;

class TableBlockCompiler {
  private readonly headerDepth = calculateTableHeaderDepth(this.block.columns);
  private readonly leafColumns = flattenColumns(this.block.columns);
  private readonly columnIdMap = createTableColumnIdMap(this.leafColumns);
  private readonly tableWidth = this.leafColumns.length;
  private readonly inlineStyle = createTableInlineStyle(this.block.border);
  private readonly footerRows = resolveFooterRows(this.block, this.leafColumns);

  /**
   * Tổng hợp tất cả data row đã render qua các group, phục vụ footer formula
   * có scope 'table'. Được cập nhật dần trong `writeDataRows`.
   */
  private readonly allRendered: RenderedDataRow[] = [];

  constructor(
    private readonly block: Extract<Block, { type: 'table' }>,
    private readonly context: CompileContext,
  ) {}

  compile(startRow: number): number {
    let row = this.writeHeader(startRow);

    if (this.block.data) {
      row = this.writeDataRows(this.block.data as Record<string, unknown>[], 0, undefined, row);
    } else if (this.block.groups) {
      row = this.writeGroups(row);
    }
    return this.writeFooterRows(row);
  }

  private writeHeader(startRow: number): number {
    for (const [offset, height] of (this.block.headerRowHeights ?? []).entries()) {
      this.context.worksheet.getRow(startRow + offset).height = height;
    }

    for (const hCell of buildHeaderMatrix(this.block.columns, this.headerDepth)) {
      const style = resolveStyle(
        hCell.style ?? this.block.headerStyle,
        undefined,
        this.context,
        'header cell',
        this.inlineStyle,
      );

      writeCell(
        this.context.worksheet,
        startRow + hCell.rowOffset,
        COL_ORIGIN + hCell.columnOffset,
        hCell.title,
        style,
      );

      if (hCell.rowSpan > 1 || hCell.colSpan > 1) {
        writeMerge(
          this.context.worksheet,
          startRow + hCell.rowOffset,
          COL_ORIGIN + hCell.columnOffset,
          startRow + hCell.rowOffset + hCell.rowSpan - 1,
          COL_ORIGIN + hCell.columnOffset + hCell.colSpan - 1,
        );
      }
    }

    for (const [offset, col] of this.leafColumns.entries()) {
      if (col.width !== undefined) {
        this.context.worksheet.getColumn(COL_ORIGIN + offset).width = col.width;
      }
      if (col.hidden !== undefined) {
        this.context.worksheet.getColumn(COL_ORIGIN + offset).hidden = col.hidden;
      }
    }

    return startRow + this.headerDepth;
  }

  private writeGroups(startRow: number): number {
    let row = startRow;
    let dataIndex = 0;

    for (const group of this.block.groups as TableGroup[]) {
      const groupRendered: RenderedDataRow[] = [];

      for (const sectionRow of group.headerRows ?? []) {
        row = this.writeSectionRow(sectionRow, groupRendered, dataIndex, row);
      }

      row = this.writeDataRows(group.data as Record<string, unknown>[], dataIndex, groupRendered, row);

      for (const sectionRow of group.footerRows ?? []) {
        row = this.writeSectionRow(sectionRow, groupRendered, dataIndex + group.data.length, row);
      }

      dataIndex += group.data.length;
    }

    return row;
  }

  private writeFooterRows(startRow: number): number {
    let row = startRow;

    for (const footerRow of this.footerRows) {
      row = this.writeSectionRow(
        { style: footerRow.style, height: footerRow.height, cells: footerRow.cells },
        this.allRendered,
        this.allRendered.length,
        row,
      );
    }

    return row;
  }

  private writeDataRows(
    rows: Record<string, unknown>[],
    firstDataIndex: number,
    groupRendered: RenderedDataRow[] | undefined,
    startRow: number,
  ): number {
    let row = startRow;

    for (const [offset, rowData] of rows.entries()) {
      const dataIndex = firstDataIndex + offset;
      const formulaCtx = this.createFormulaContext(row);

      const zebraStyle = dataIndex % 2 === 0 ? this.block.oddRowStyle : this.block.evenRowStyle;

      for (const [colOffset, col] of this.leafColumns.entries()) {
        const rawValue = resolveTableCellValue(rowData, col);
        assertTableCellValue(rawValue);
        const compiled = FormulaCompiler.compileCellContent(rawValue, formulaCtx);

        const style = resolveStyle(
          col.bodyStyle ?? col.style ?? zebraStyle ?? this.block.bodyStyle,
          col.styleResolver?.(compiled.value, rowData, dataIndex),
          this.context,
          `table column "${String(col.id ?? col.title)}"`,
          this.inlineStyle,
        );

        const valueOrFormula = FormulaCompiler.isFormulaDefinition(rawValue)
          ? { formula: compiled.formula }
          : compiled.value;

        writeCell(this.context.worksheet, row, COL_ORIGIN + colOffset, valueOrFormula, style);
      }

      if (this.block.bodyRowHeight !== undefined) {
        this.context.worksheet.getRow(row).height = this.block.bodyRowHeight;
      }

      if (this.block.rowHidden?.(rowData, dataIndex) === true) {
        this.context.worksheet.getRow(row).hidden = true;
      }

      const rendered: RenderedDataRow = { data: rowData, row };
      groupRendered?.push(rendered);
      this.allRendered.push(rendered);
      row++;
    }

    return row;
  }

  private writeSectionRow(
    sectionRow: TableSectionRow,
    currentRendered: RenderedDataRow[],
    dataIndex: number,
    row: number,
  ): number {
    const occupied = new Set<number>();
    const placements = this.createSectionCellPlacements(sectionRow, occupied, row);
    const formulaCtx = this.createFormulaContext(row, {
      group: currentRendered,
      table: this.allRendered,
    });

    if (sectionRow.height !== undefined) {
      this.context.worksheet.getRow(row).height = sectionRow.height;
    }
    if (sectionRow.hidden !== undefined) {
      this.context.worksheet.getRow(row).hidden = sectionRow.hidden;
    }

    for (const { cell, cellIndex, columnOffset, colSpan } of placements) {
      const rawValue = resolveSectionCellValue(cell, {
        rows: currentRendered.map((r) => r.data),
        allRows: this.allRendered.map((r) => r.data),
        dataIndex,
        rowIndex: row,
      });

      assertTableCellValue(rawValue);
      const compiled = FormulaCompiler.compileCellContent(rawValue, formulaCtx);

      const style = resolveStyle(
        cell.style ?? sectionRow.style,
        cell.styleResolver?.(compiled.value),
        this.context,
        `table section cell ${cellIndex + 1}`,
        this.inlineStyle,
      );

      const valueOrFormula = FormulaCompiler.isFormulaDefinition(rawValue)
        ? { formula: compiled.formula }
        : compiled.value;

      writeCell(this.context.worksheet, row, COL_ORIGIN + columnOffset, valueOrFormula, style);

      if (colSpan > 1) {
        writeMerge(
          this.context.worksheet,
          row,
          COL_ORIGIN + columnOffset,
          row,
          COL_ORIGIN + columnOffset + colSpan - 1,
        );
      }
    }

    this.fillSectionRowStyle(sectionRow, occupied, row);
    return row + 1;
  }

  private createSectionCellPlacements(
    sectionRow: TableSectionRow,
    occupied: Set<number>,
    row: number,
  ): SectionCellPlacement[] {
    const placements: SectionCellPlacement[] = [];

    for (const [cellIndex, cell] of sectionRow.cells.entries()) {
      const colOffset = resolveSectionCellColumnOffset(cell, cellIndex, this.columnIdMap, occupied);
      const colSpan = resolveSectionCellColSpan(cell, colOffset, this.tableWidth);

      if (cell.id) {
        this.context.registry.register(
          this.context.sheet.id,
          this.context.sheet.name,
          cell.id,
          row,
          COL_ORIGIN + colOffset,
        );
      }

      for (let o = colOffset; o < colOffset + colSpan; o++) {
        if (occupied.has(o)) throw new CompileError('Table section row cells must not overlap.');
        occupied.add(o);
      }

      placements.push({ cell, cellIndex, columnOffset: colOffset, colSpan });
    }

    return placements;
  }

  private createFormulaContext(row: number, scopedRows?: Record<'group' | 'table', RenderedDataRow[]>) {
    return FormulaResolver.createTableContext({
      columnIdMap: this.columnIdMap,
      row,
      firstColumn: COL_ORIGIN,
      registry: this.context.registry,
      currentSheetId: this.context.sheet.id,
      scopedRows,
    });
  }

  private fillSectionRowStyle(sectionRow: TableSectionRow, occupied: Set<number>, row: number): void {
    if (!sectionRow.style) return;

    for (let o = 0; o < this.tableWidth; o++) {
      if (!occupied.has(o)) {
        const style = resolveStyle(sectionRow.style, undefined, this.context, 'section filler', this.inlineStyle);
        writeCell(this.context.worksheet, row, COL_ORIGIN + o, null, style);
      }
    }
  }
}
