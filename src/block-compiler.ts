import type ExcelJS from 'exceljs';
import { AddressRegistry } from './address-registry';
import { CompileError, ReportEngineError } from './errors';
import { createGridFormulaContext, createTableFormulaContext, type RenderedDataRow } from './formula-context';
import { compileCellContent, isFormulaDefinition } from './formula-engine';
import { writeCell, writeMerge } from './helpers/exceljs';
import { mergeCellStyles } from './helpers/style';
import {
  assertGridCellDoesNotOverlap,
  buildHeaderMatrix,
  calculateTableHeaderDepth,
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
import type {
  Block,
  CellContent,
  CellStyleDefinition,
  GridCell,
  SheetDefinition,
  StyleRegistry,
  StyleValue,
  TableFooterRow,
  TableGroup,
  TableLeafColumn,
  TableSectionCell,
  TableSectionRow,
  WorkbookDefinition,
} from './types';
import { interpolateCellValue, interpolateVariables, type VariableScope } from './variable-engine';

export interface CompileContext {
  readonly workbook: WorkbookDefinition;
  readonly sheet: SheetDefinition;
  readonly sheetColumnCount: number;
  readonly variables: VariableScope;
  readonly registry: AddressRegistry;
  readonly worksheet: ExcelJS.Worksheet;
  readonly styleConfig: {
    defaultStyle?: CellStyleDefinition;
    styles?: StyleRegistry;
  };
}

// ─── Public entry point ──────────────────────────────────────────────────────

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

// ─── Grid ─────────────────────────────────────────────────────────────────────

interface GridPlacement {
  cell: GridCell;
  row: number;
  column: number;
  rowSpan: number;
  colSpan: number;
}

function compileGridBlock(block: Extract<Block, { type: 'grid' }>, context: CompileContext, startRow: number): number {
  const occupied = new Set<string>();
  const placements: GridPlacement[] = [];
  const variables = blockVariables(context, block);
  let rowExtent = block.rows.length;

  // Pass 1: layout — compute positions and register IDs into registry
  for (const [rowOffset, gridRow] of block.rows.entries()) {
    const absRow = startRow + rowOffset;

    if (gridRow.height !== undefined) {
      context.worksheet.getRow(absRow).height = gridRow.height;
    }

    let colOffset = 0;

    for (const cell of gridRow.cells) {
      while (occupied.has(gridOccupancyKey(rowOffset, colOffset))) colOffset++;

      const colSpan = resolveColSpan(cell.colSpan, colOffset, context.sheetColumnCount);
      const rowSpan = cell.rowSpan ?? 1;
      assertGridCellDoesNotOverlap(occupied, rowOffset, colOffset, rowSpan, colSpan);

      const absCol = 1 + colOffset;

      if (cell.id) {
        context.registry.register(context.sheet.id, context.sheet.name, cell.id, absRow, absCol);
      }

      placements.push({ cell, row: absRow, column: absCol, rowSpan, colSpan });
      markGridOccupied(occupied, rowOffset, colOffset, rowSpan, colSpan);
      rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
      colOffset += colSpan;
    }
  }

  // Pass 2: render — all IDs registered, now compile formulas
  const formulaCtx = createGridFormulaContext(context.registry, context.sheet.id);

  for (const p of placements) {
    const compiled = compileCellContent(interpolateCellValue(p.cell.value, variables), formulaCtx);

    const style = resolveStyle(
      p.cell.style,
      p.cell.styleResolver?.(compiled.value),
      context,
      `grid cell "${p.cell.id ?? `${p.row}:${p.column}`}"`,
    );

    const valueOrFormula = isFormulaDefinition(p.cell.value)
      ? { formula: compiled.formula, result: p.cell.formulaResult }
      : compiled.value;

    writeCell(context.worksheet, p.row, p.column, valueOrFormula, style);

    if (p.cell.width !== undefined) {
      context.worksheet.getColumn(p.column).width = p.cell.width;
    }

    if (p.rowSpan > 1 || p.colSpan > 1) {
      writeMerge(context.worksheet, p.row, p.column, p.row + p.rowSpan - 1, p.column + p.colSpan - 1);
    }
  }

  return startRow + rowExtent;
}

// ─── Table ────────────────────────────────────────────────────────────────────

function compileTableBlock(
  block: Extract<Block, { type: 'table' }>,
  context: CompileContext,
  startRow: number,
): number {
  return new TableBlockCompiler(block, context).compile(startRow);
}

// ─── Table helpers ────────────────────────────────────────────────────────────

interface SectionCellPlacement {
  cell: TableSectionCell;
  cellIndex: number;
  columnOffset: number;
  colSpan: number;
}

class TableBlockCompiler {
  private readonly variables = blockVariables(this.context, this.block);
  private readonly headerDepth = calculateTableHeaderDepth(this.block.columns);
  private readonly leafColumns = flattenColumns(this.block.columns);
  private readonly columnIdMap = createTableColumnIdMap(this.leafColumns);
  private readonly tableWidth = this.leafColumns.length;
  private readonly inlineStyle = createTableInlineStyle(this.block.border);
  private readonly footerRows = resolveFooterRows(this.block, this.leafColumns);

  /**
   * Tracks all rendered data rows across groups for footer section formulas
   * (scope: 'allRows'). Built incrementally during writeDataRows.
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
        1 + hCell.columnOffset,
        interpolateVariables(hCell.title, this.variables),
        style,
      );

      if (hCell.rowSpan > 1 || hCell.colSpan > 1) {
        writeMerge(
          this.context.worksheet,
          startRow + hCell.rowOffset,
          1 + hCell.columnOffset,
          startRow + hCell.rowOffset + hCell.rowSpan - 1,
          1 + hCell.columnOffset + hCell.colSpan - 1,
        );
      }
    }

    for (const [offset, col] of this.leafColumns.entries()) {
      if (col.width !== undefined) {
        this.context.worksheet.getColumn(1 + offset).width = col.width;
      }
      if (col.hidden !== undefined) {
        this.context.worksheet.getColumn(1 + offset).hidden = col.hidden;
      }
    }

    return startRow + this.headerDepth;
  }

  private writeGroups(startRow: number): number {
    let row = startRow;
    let dataIndex = 0;

    for (const group of (this.block.groups ?? []) as TableGroup[]) {
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
      const bandStyle = dataIndex % 2 === 0 ? this.block.oddRowStyle : this.block.evenRowStyle;

      for (const [colOffset, col] of this.leafColumns.entries()) {
        const rawValue = resolveTableCellValue(rowData, col);
        assertTableCellValue(rawValue);
        const compiled = compileCellContent(interpolateCellValue(rawValue, this.variables), formulaCtx);

        const style = resolveStyle(
          col.bodyStyle ?? col.style ?? bandStyle ?? this.block.bodyStyle,
          col.styleResolver?.(compiled.value, rowData, dataIndex),
          this.context,
          `table column "${String(col.id ?? col.title)}"`,
          this.inlineStyle,
        );

        const valueOrFormula = isFormulaDefinition(rawValue) ? { formula: compiled.formula } : compiled.value;

        writeCell(this.context.worksheet, row, 1 + colOffset, valueOrFormula, style);
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
      currentRows: currentRendered,
      allRows: this.allRendered,
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
      const compiled = compileCellContent(interpolateCellValue(rawValue, this.variables), formulaCtx);

      const style = resolveStyle(
        cell.style ?? sectionRow.style,
        cell.styleResolver?.(compiled.value),
        this.context,
        `table section cell ${cellIndex + 1}`,
        this.inlineStyle,
      );

      const valueOrFormula = isFormulaDefinition(rawValue) ? { formula: compiled.formula } : compiled.value;

      writeCell(this.context.worksheet, row, 1 + columnOffset, valueOrFormula, style);

      if (colSpan > 1) {
        writeMerge(this.context.worksheet, row, 1 + columnOffset, row, 1 + columnOffset + colSpan - 1);
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
        this.context.registry.register(this.context.sheet.id, this.context.sheet.name, cell.id, row, 1 + colOffset);
      }

      for (let o = colOffset; o < colOffset + colSpan; o++) {
        if (occupied.has(o)) throw new CompileError('Table section row cells must not overlap.');
        occupied.add(o);
      }

      placements.push({ cell, cellIndex, columnOffset: colOffset, colSpan });
    }

    return placements;
  }

  private createFormulaContext(row: number, scopedRows?: Record<'currentRows' | 'allRows', RenderedDataRow[]>) {
    return createTableFormulaContext({
      columnIdMap: this.columnIdMap,
      row,
      firstColumn: 1,
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
        writeCell(this.context.worksheet, row, 1 + o, null, style);
      }
    }
  }
}

// ─── Helpers ──────────────────────────────────────────────────────────────────

function resolveFooterRows(block: Extract<Block, { type: 'table' }>, leafColumns: TableLeafColumn[]): TableFooterRow[] {
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
  inlineStyle?: CellStyleDefinition,
): CellStyleDefinition | undefined {
  const styleVal = dynamicStyle !== undefined ? dynamicStyle : staticStyle;

  let baseStyle: CellStyleDefinition | undefined;
  if (typeof styleVal === 'string') {
    const registryStyle = context.styleConfig.styles?.[styleVal];
    if (!registryStyle) {
      throw new ReportEngineError(`${label} returned unknown style "${styleVal}".`);
    }
    baseStyle = registryStyle;
  } else {
    baseStyle = styleVal;
  }

  return mergeCellStyles(mergeCellStyles(context.styleConfig.defaultStyle, baseStyle), inlineStyle);
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
