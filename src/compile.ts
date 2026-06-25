import { ReportEngineError } from './errors';
import type { Block, SheetDefinition, TableSectionRow, WorkbookDefinition } from './types';
import { validateWorkbookDefinition } from './validation';
import { compileBlock } from './block-compiler';
import { type CellAddress, createFormulaId } from './formula-engine';
import { LayoutCursor } from './layout-cursor';
import type { RenderPlan } from './render-plan';
import { RenderPlanBuilder } from './render-plan-builder';
import type { RenderContext } from './variable-engine';
import {
  assertNever,
  calculateTableHeaderDepth,
  createTableColumnIdMap,
  createSheetColumnCounts,
  flattenColumns,
  gridOccupancyKey,
  markGridOccupied,
  resolveColSpan,
  resolveSectionCellColumnOffset,
  resolveSectionCellColSpan,
} from './helpers/utils';

export interface CompileWorkbookOptions {
  context?: RenderContext;
}

export function compileWorkbookToRenderPlan(
  workbook: WorkbookDefinition,
  options: CompileWorkbookOptions = {},
): RenderPlan {
  validateWorkbookDefinition(workbook);

  const sheetColumnCounts = createSheetColumnCounts(workbook);
  const formulaIds = collectWorkbookFormulaIds(workbook, sheetColumnCounts);
  const workbookContext = { ...(workbook.context ?? {}), ...(options.context ?? {}) };
  const builder = new RenderPlanBuilder({
    metadata: workbook.metadata,
    defaultStyle: workbook.defaultStyle,
    styles: workbook.styles,
  });

  for (const sheet of workbook.sheets) {
    builder.addSheet(sheet.id, sheet.name, {
      views: sheet.freezePane
        ? [{ state: 'frozen', xSplit: sheet.freezePane.columns, ySplit: sheet.freezePane.rows }]
        : undefined,
    });
  }

  for (const sheet of workbook.sheets) {
    const context = {
      workbook,
      sheet,
      styles: workbook.styles,
      variables: { workbook: workbookContext, sheet: sheet.context },
      formulaIds,
      sheetColumnCount: sheetColumnCounts.get(sheet.id) ?? 1,
    };
    const cursor = new LayoutCursor();

    for (const block of sheet.blocks) {
      compileBlock(block, context, cursor, builder);
    }
  }

  return builder.build();
}

function collectWorkbookFormulaIds(
  workbook: WorkbookDefinition,
  sheetColumnCounts: ReadonlyMap<string, number>,
): Map<string, CellAddress> {
  const formulaIds = new Map<string, CellAddress>();

  for (const sheet of workbook.sheets) {
    const cursor = new LayoutCursor();

    for (const block of sheet.blocks) {
      switch (block.type) {
        case 'title':
        case 'text':
          cursor.advanceRows();
          break;

        case 'spacer':
        case 'divider':
          cursor.advanceRows(block.rows ?? 1);
          break;

        case 'grid':
          cursor.advanceRows(
            collectGridFormulaIds(formulaIds, sheet, block, cursor, sheetColumnCounts.get(sheet.id) ?? 1),
          );
          break;

        case 'table':
        case 'table-groups':
          cursor.advanceRows(collectTableFormulaIds(formulaIds, sheet, block, cursor));
          break;

        default:
          assertNever(block);
      }
    }
  }

  return formulaIds;
}

function collectGridFormulaIds(
  formulaIds: Map<string, CellAddress>,
  sheet: SheetDefinition,
  block: Extract<Block, { type: 'grid' }>,
  cursor: LayoutCursor,
  sheetColumnCount: number,
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
      const colSpan = resolveColSpan(cell.colSpan, columnOffset, sheetColumnCount);

      if (cell.id) {
        registerWorkbookFormulaId(formulaIds, sheet, cell.id, cursor.row + rowOffset, cursor.column + columnOffset);
      }

      markGridOccupied(occupied, rowOffset, columnOffset, rowSpan, colSpan);
      rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
      columnOffset += colSpan;
    }
  }

  return rowExtent;
}

function registerWorkbookFormulaId(
  formulaIds: Map<string, CellAddress>,
  sheet: SheetDefinition,
  id: string,
  row: number,
  column: number,
): void {
  const registryId = createFormulaId(sheet.id, id);

  if (formulaIds.has(registryId)) {
    throw new ReportEngineError(`ID formula cell "${id}" bị trùng lặp trong sheet "${sheet.id}".`);
  }

  formulaIds.set(registryId, {
    row,
    column,
    sheetId: sheet.id,
    sheetName: sheet.name,
  });
}

function collectTableFormulaIds(
  formulaIds: Map<string, CellAddress>,
  sheet: SheetDefinition,
  block: Extract<Block, { type: 'table' | 'table-groups' }>,
  cursor: LayoutCursor,
): number {
  const headerDepth = calculateTableHeaderDepth(block.columns);
  const leafColumns = flattenColumns(block.columns);
  const columnIdMap = createTableColumnIdMap(leafColumns);
  const tableWidth = leafColumns.length;

  let rowOffset = headerDepth;

  if (block.type === 'table') {
    rowOffset += block.data.length;
  } else {
    for (const group of block.groups) {
      for (const sectionRow of group.headerRows ?? []) {
        collectTableSectionFormulaIds(
          formulaIds,
          sheet,
          sectionRow,
          cursor.row + rowOffset,
          cursor.column,
          columnIdMap,
          tableWidth,
        );
        rowOffset += 1;
      }

      rowOffset += group.data.length;

      for (const sectionRow of group.footerRows ?? []) {
        collectTableSectionFormulaIds(
          formulaIds,
          sheet,
          sectionRow,
          cursor.row + rowOffset,
          cursor.column,
          columnIdMap,
          tableWidth,
        );
        rowOffset += 1;
      }
    }
  }

  for (const footerRow of block.footerRows ?? []) {
    collectTableSectionFormulaIds(
      formulaIds,
      sheet,
      { style: footerRow.style, height: footerRow.height, cells: footerRow.cells },
      cursor.row + rowOffset,
      cursor.column,
      columnIdMap,
      tableWidth,
    );
    rowOffset += 1;
  }

  if (!block.footerRows && flattenColumns(block.columns).some((col) => col.summary !== undefined)) {
    rowOffset += 1;
  }

  return rowOffset;
}

function collectTableSectionFormulaIds(
  formulaIds: Map<string, CellAddress>,
  sheet: SheetDefinition,
  sectionRow: TableSectionRow,
  row: number,
  firstColumn: number,
  columnIdMap: Map<string, number>,
  tableWidth: number,
): void {
  const occupiedColumns = new Set<number>();

  for (const [cellIndex, cell] of sectionRow.cells.entries()) {
    const columnOffset = resolveSectionCellColumnOffset(cell, cellIndex, columnIdMap, occupiedColumns);
    const colSpan = resolveSectionCellColSpan(cell, columnOffset, tableWidth);

    if (cell.id) {
      registerWorkbookFormulaId(formulaIds, sheet, cell.id, row, firstColumn + columnOffset);
    }

    for (let offset = columnOffset; offset < columnOffset + colSpan; offset += 1) {
      occupiedColumns.add(offset);
    }
  }
}
