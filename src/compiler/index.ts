import { ReportEngineError } from '../core/errors';
import type {
  Block,
  SheetDefinition,
  TableFooterRow,
  TableLeafColumn,
  TableSectionRow,
  WorkbookDefinition,
} from '../core/types';
import { validateWorkbookDefinition } from '../core/validation';
import { compileBlock } from './block-compiler';
import { type CellAddress, createFormulaId } from './formula-engine';
import { LayoutCursor } from './layout-cursor';
import type { RenderPlan, ResolvedNamedRange } from './render-plan';
import { RenderPlanBuilder } from './render-plan-builder';
import type { RenderContext } from './variable-engine';
import { calculateTableHeaderDepth, flattenColumns } from '../helpers/utils'; // ─── Public Types ────────────────────────────────────────────────────────────

// ─── Public Types ────────────────────────────────────────────────────────────

export interface CompileWorkbookOptions {
  /** Context bên ngoài được merge vào workbook.context trước khi compile. */
  context?: RenderContext;
}

// ─── Internal Types ───────────────────────────────────────────────────────────

// ─── Entry Point ─────────────────────────────────────────────────────────────

/**
 * Biên dịch toàn bộ WorkbookDefinition thành một RenderPlan sẵn sàng để render ra Excel.
 *
 * Các bước thực hiện:
 *  1. Validate cấu trúc workbook.
 *  2. Xây dựng đồ thị phụ thuộc giữa các sheet (dùng cho sort hoặc cảnh báo vòng lặp sau này).
 *  3. Thu thập địa chỉ ô (row/col) của tất cả formula cell có ID.
 *  4. Giải quyết named ranges thành tọa độ thực.
 *  5. Compile từng block trên từng sheet và ghi vào RenderPlanBuilder.
 */
export function compileWorkbookToRenderPlan(
  workbook: WorkbookDefinition,
  options: CompileWorkbookOptions = {},
): RenderPlan {
  // Bước 1: validate toàn bộ cấu trúc workbook trước khi làm bất cứ điều gì.
  validateWorkbookDefinition(workbook);

  // Bước 3: duyệt tất cả block, tính row/col thực tế của từng formula cell có ID.
  const formulaIds = collectWorkbookFormulaIds(workbook);

  // Bước 4: ánh xạ named ranges (startId/endId) → tọa độ ô thực (row/col).
  const namedRanges = resolveNamedRanges(workbook, formulaIds);
  const namedRangeNames = new Set((namedRanges ?? []).map((range) => range.name));

  // Merge context: workbook-level context được override bởi options.context.
  const workbookContext = { ...(workbook.context ?? {}), ...(options.context ?? {}) };

  // Bước 5: khởi tạo builder và điền dữ liệu.
  const builder = new RenderPlanBuilder({
    metadata: workbook.metadata,
    defaultStyle: workbook.defaultStyle,
    styles: workbook.styles,
    namedRanges,
  });

  // Đăng ký tất cả sheet (tên + freeze pane) trước khi ghi nội dung.
  for (const sheet of workbook.sheets) {
    builder.addSheet(sheet.id, sheet.name, {
      views: sheet.freezePane
        ? [{ state: 'frozen', xSplit: sheet.freezePane.columns, ySplit: sheet.freezePane.rows }]
        : undefined,
    });
  }

  // Compile từng sheet: mỗi sheet có cursor riêng theo dõi hàng hiện tại.
  for (const sheet of workbook.sheets) {
    const context = {
      workbook,
      sheet,
      styles: workbook.styles,
      variables: { workbook: workbookContext, sheet: sheet.context },
      formulaIds,
      namedRangeNames,
    };
    const cursor = new LayoutCursor();

    // expandBlocks xử lý repeat block trước, trả về danh sách block phẳng.
    for (const block of sheet.blocks) {
      compileBlock(block, context, cursor, builder);
    }
  }

  return builder.build();
}

// ─── Formula ID Collection ────────────────────────────────────────────────────

/**
 * Duyệt tất cả block của workbook, tính toán địa chỉ ô tuyệt đối (row, column)
 * của mỗi formula cell có thuộc tính `id`.
 *
 * Địa chỉ này sau đó được dùng để:
 *  - Compile cross-cell reference (=Sheet1!B5).
 *  - Giải quyết named ranges.
 */
function collectWorkbookFormulaIds(workbook: WorkbookDefinition): Map<string, CellAddress> {
  const formulaIds = new Map<string, CellAddress>();

  for (const sheet of workbook.sheets) {
    const cursor = new LayoutCursor();

    for (const block of sheet.blocks) {
      switch (block.type) {
        // Các block không có formula ID — chỉ advance cursor.
        case 'title':
        case 'text':
          cursor.advanceRows();
          break;

        case 'spacer':
        case 'divider':
          cursor.advanceRows(block.rows ?? 1);
          break;

        // Grid: từng cell có thể có ID, cần tính row/col chính xác kể cả rowSpan/colSpan.
        case 'grid':
          cursor.advanceRows(collectGridFormulaIds(formulaIds, sheet, block, cursor));
          break;

        // Table: header + data rows + section rows + footer rows.
        case 'table':
          cursor.advanceRows(collectTableFormulaIds(formulaIds, sheet, block, cursor));
          break;

        default:
          assertNever(block);
      }
    }
  }

  return formulaIds;
}

/**
 * Thu thập formula ID từ grid block.
 * Grid có thể có rowSpan/colSpan nên cần theo dõi ô nào đã bị chiếm (occupied set).
 *
 * @returns Số hàng mà grid block này chiếm (để advance cursor).
 */
function collectGridFormulaIds(
  formulaIds: Map<string, CellAddress>,
  sheet: SheetDefinition,
  block: Extract<Block, { type: 'grid' }>,
  cursor: LayoutCursor,
): number {
  const occupied = new Set<string>(); // Tập các ô đã bị chiếm bởi rowSpan/colSpan.
  let rowExtent = block.rows.length; // Số hàng tối thiểu = số hàng định nghĩa trong block.

  for (const [rowOffset, gridRow] of block.rows.entries()) {
    let columnOffset = 0;

    for (const cell of gridRow.cells) {
      // Bỏ qua ô đã bị chiếm bởi cell phía trên (do rowSpan).
      while (occupied.has(gridOccupancyKey(rowOffset, columnOffset))) {
        columnOffset += 1;
      }

      const rowSpan = cell.rowSpan ?? 1;
      const colSpan = cell.colSpan ?? 1;

      if (cell.id) {
        registerWorkbookFormulaId(formulaIds, sheet, cell.id, cursor.row + rowOffset, cursor.column + columnOffset);
      }

      // Đánh dấu toàn bộ vùng cell này là đã chiếm.
      markGridOccupied(occupied, rowOffset, columnOffset, rowSpan, colSpan);

      // Cập nhật số hàng thực tế nếu rowSpan vượt ra ngoài.
      rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
      columnOffset += colSpan;
    }
  }

  return rowExtent;
}

/**
 * Đăng ký một formula cell ID vào global map.
 * Ném lỗi nếu ID bị trùng trong cùng một sheet.
 */
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

// ─── Named Range Resolution ───────────────────────────────────────────────────

/**
 * Chuyển named range definitions (dùng startId/endId) thành tọa độ ô thực tế.
 * Named range phải nằm trong cùng một sheet và end phải sau start.
 *
 * @returns undefined nếu workbook không có named range nào.
 */
function resolveNamedRanges(
  workbook: WorkbookDefinition,
  formulaIds: Map<string, CellAddress>,
): ResolvedNamedRange[] | undefined {
  if (!workbook.namedRanges?.length) return undefined;

  return workbook.namedRanges.map((range) => {
    const start = formulaIds.get(createFormulaId(range.sheetId, range.startId));
    const end = formulaIds.get(createFormulaId(range.sheetId, range.endId));

    if (!start) {
      throw new ReportEngineError(`Named range "${range.name}" references unknown startId "${range.startId}".`);
    }
    if (!end) {
      throw new ReportEngineError(`Named range "${range.name}" references unknown endId "${range.endId}".`);
    }
    if (end.row < start.row || end.column < start.column) {
      throw new ReportEngineError(`Named range "${range.name}" endId must resolve after startId.`);
    }

    const sheet = workbook.sheets.find((s) => s.id === range.sheetId);
    if (!sheet) {
      throw new ReportEngineError(`Named range "${range.name}" references unknown sheetId "${range.sheetId}".`);
    }

    return {
      name: range.name,
      sheetName: sheet.name,
      startRow: start.row,
      startColumn: start.column,
      endRow: end.row,
      endColumn: end.column,
    };
  });
}

// ─── Table Formula ID Collection ─────────────────────────────────────────────

/**
 * Tính địa chỉ ô của mỗi formula cell có ID trong table block.
 * Thứ tự hàng: titleRows → header rows → data rows (kể cả section rows) → footer rows → summary row.
 *
 * @returns Số hàng mà table block chiếm (để advance cursor).
 */
function collectTableFormulaIds(
  formulaIds: Map<string, CellAddress>,
  sheet: SheetDefinition,
  block: Extract<Block, { type: 'table' }>,
  cursor: LayoutCursor,
): number {
  const headerDepth = calculateTableHeaderDepth(block.columns);
  const columnIdMap = createTableColumnIdMap(flattenColumns(block.columns));

  // Offset bắt đầu sau title rows và header rows.
  let rowOffset = (block.titleRows?.length ?? 0) + headerDepth;

  for (const item of block.data as unknown[]) {
    if (isTableSectionRow(item)) {
      collectTableSectionFormulaIds(formulaIds, sheet, item, cursor.row + rowOffset, cursor.column, columnIdMap);
    }
    rowOffset += 1;
  }

  // Footer rows có ID riêng.
  for (const footerRow of block.footerRows ?? []) {
    collectTableFooterFormulaIds(formulaIds, sheet, footerRow, cursor.row + rowOffset, cursor.column, columnIdMap);
    rowOffset += 1;
  }

  // Summary row tự động (khi không có footerRows nhưng có column summary).
  if (!block.footerRows && flattenColumns(block.columns).some((col) => col.summary !== undefined)) {
    rowOffset += 1;
  }

  return rowOffset;
}

/** Delegate footer row vào collectTableSectionFormulaIds (footer có cùng cấu trúc với section). */
function collectTableFooterFormulaIds(
  formulaIds: Map<string, CellAddress>,
  sheet: SheetDefinition,
  footerRow: TableFooterRow,
  row: number,
  firstColumn: number,
  columnIdMap: Map<string, number>,
): void {
  collectTableSectionFormulaIds(
    formulaIds,
    sheet,
    { type: 'section', style: footerRow.style, height: footerRow.height, cells: footerRow.cells },
    row,
    firstColumn,
    columnIdMap,
  );
}

/**
 * Thu thập formula ID từ một section row (hoặc footer row).
 * Cell trong section có thể tham chiếu column theo: thứ tự (mặc định), column index, hoặc columnId.
 */
function collectTableSectionFormulaIds(
  formulaIds: Map<string, CellAddress>,
  sheet: SheetDefinition,
  sectionRow: TableSectionRow,
  row: number,
  firstColumn: number,
  columnIdMap: Map<string, number>,
): void {
  const occupiedColumns = new Set<number>();

  for (const [cellIndex, cell] of sectionRow.cells.entries()) {
    const columnOffset = resolveTableSectionCellColumnOffset(cell, cellIndex, columnIdMap, occupiedColumns);
    const colSpan = cell.colSpan === 'remaining' ? 1 : (cell.colSpan ?? 1);

    if (cell.id) {
      registerWorkbookFormulaId(formulaIds, sheet, cell.id, row, firstColumn + columnOffset);
    }

    // Đánh dấu tất cả cột bị chiếm bởi cell này.
    for (let offset = columnOffset; offset < columnOffset + colSpan; offset += 1) {
      occupiedColumns.add(offset);
    }
  }
}

/**
 * Xác định columnOffset cho một cell trong section/footer row.
 * Ưu tiên: cell.column (1-based) > cell.columnId > tự động (tìm cột trống tiếp theo).
 */
function resolveTableSectionCellColumnOffset(
  cell: TableSectionRow['cells'][number],
  cellIndex: number,
  columnIdMap: Map<string, number>,
  occupiedColumns: Set<number>,
): number {
  if (cell.column !== undefined) return cell.column - 1; // column là 1-based.

  if (cell.columnId !== undefined) {
    const offset = columnIdMap.get(cell.columnId);
    if (offset === undefined) {
      throw new ReportEngineError(`Table section row references unknown columnId "${cell.columnId}".`);
    }
    return offset;
  }

  // Tự động: bắt đầu từ 0 (cell đầu tiên) hoặc max(occupied)+1, rồi tìm cột chưa chiếm.
  let offset = cellIndex === 0 ? 0 : Math.max(...occupiedColumns) + 1;
  while (occupiedColumns.has(offset)) offset += 1;
  return offset;
}

// ─── Table Helpers ────────────────────────────────────────────────────────────

/**
 * Tạo Map: columnId → columnOffset (0-based) từ danh sách leaf column.
 * Dùng để resolve `cell.columnId` trong section/footer row.
 */
function createTableColumnIdMap(columns: TableLeafColumn[]): Map<string, number> {
  const idMap = new Map<string, number>();

  for (const [columnOffset, column] of columns.entries()) {
    if (!column.id) continue;

    const key = String(column.id);
    if (idMap.has(key)) {
      throw new ReportEngineError(`Duplicate formula cell id "${key}".`);
    }
    idMap.set(key, columnOffset);
  }

  return idMap;
}

// ─── Grid Utilities ───────────────────────────────────────────────────────────

/**
 * Đánh dấu vùng [rowOffset, rowOffset+rowSpan) × [columnOffset, columnOffset+colSpan)
 * là đã bị chiếm trong grid, ngăn các cell khác ghi đè lên.
 */
function markGridOccupied(
  occupied: Set<string>,
  rowOffset: number,
  columnOffset: number,
  rowSpan: number,
  colSpan: number,
): void {
  for (let row = rowOffset; row < rowOffset + rowSpan; row += 1) {
    for (let col = columnOffset; col < columnOffset + colSpan; col += 1) {
      occupied.add(gridOccupancyKey(row, col));
    }
  }
}

/** Tạo key duy nhất cho một ô trong grid occupancy tracking. */
function gridOccupancyKey(rowOffset: number, columnOffset: number): string {
  return `${rowOffset}:${columnOffset}`;
}

// ─── Type Guards & Utilities ──────────────────────────────────────────────────

/** Kiểm tra value có phải TableSectionRow không. */
function isTableSectionRow(value: unknown): value is TableSectionRow {
  return (
    typeof value === 'object' && value !== null && 'type' in value && (value as { type: unknown }).type === 'section'
  );
}

/** Xử lý trường hợp không bao giờ xảy ra (exhaustive switch). */
function assertNever(value: never): never {
  throw new ReportEngineError(`Unsupported block type "${(value as Block).type}".`);
}

// ─── Re-exports ───────────────────────────────────────────────────────────────

export { compileBlock } from './block-compiler';
export {
  compileCellContent,
  compileFormula,
  createFormulaCompileContext,
  createFormulaId,
  formatCellAddress,
  formatCellReference,
  isFormulaDefinition,
} from './formula-engine';
export { LayoutCursor } from './layout-cursor';
export { assertMergeDoesNotOverlap, normalizeMergeRange } from './merge-engine';
export { RenderPlanBuilder } from './render-plan-builder';
export { interpolateCellValue, interpolateVariables, resolvePath } from './variable-engine';

export type { BlockCompiler, BlockCompilerRegistry, SheetContext } from './block-compiler';
export type { CellAddress, FormulaCompileContext, FormulaCompileContextOptions } from './formula-engine';
export type { MergeRange, NormalizedMergeRange } from './merge-engine';
export type { RenderContext, VariableScope } from './variable-engine';
