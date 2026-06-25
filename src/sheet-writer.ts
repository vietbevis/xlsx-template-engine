import { ReportEngineError } from './errors';
import { assertPositiveInteger } from './helpers/common';
import { cloneStylePart } from './helpers/style';
import { assertMergeDoesNotOverlap, normalizeMergeRange, type MergeRange } from './merge-engine';
import type {
  RenderCell,
  RenderColumnVisibility,
  RenderColumnWidth,
  RenderMergeRange,
  RenderPlanSheet,
  RenderRowHeight,
  RenderRowVisibility,
  RenderSheetView,
} from './render-plan';

/**
 * Writer ghi trực tiếp vào một `RenderPlanSheet` mà không qua lớp builder trung gian.
 *
 * - Không clone hai lần (addCell + build()).
 * - Không lưu rowMap song song với rows array — dùng Map<row, index> nhỏ gọn hơn.
 * - scoped theo sheet: mỗi sheet có một SheetWriter riêng → loại bỏ việc
 *   truyền `sheetId` vào mọi lời gọi.
 */
export class SheetWriter {
  private readonly sheet: RenderPlanSheet;
  private readonly rowIndex = new Map<number, number>(); // rowNumber → index in sheet.rows
  private readonly mergeRanges: MergeRange[] = [];

  constructor(id: string, name: string, views?: RenderSheetView[]) {
    this.sheet = {
      id,
      name,
      views: views ? views.map((v) => ({ ...v })) : undefined,
      rows: [],
      merges: [],
      columnWidths: [],
      columnVisibility: [],
      rowHeights: [],
      rowVisibility: [],
    };
  }

  addCell(cell: RenderCell): void {
    assertPositiveInteger(cell.row, 'cell row');
    assertPositiveInteger(cell.column, 'cell column');
    this.getOrCreateRow(cell.row).cells.push(cloneCell(cell));
  }

  addMerge(range: RenderMergeRange): void {
    const normalized = normalizeMergeRange(this.sheet.id, range);

    if (normalized.type === 'skip-single-cell') {
      return;
    }

    assertMergeDoesNotOverlap(normalized.range, this.mergeRanges);
    this.mergeRanges.push(normalized.range);
    this.sheet.merges.push({
      startRow: normalized.range.startRow,
      startColumn: normalized.range.startColumn,
      endRow: normalized.range.endRow,
      endColumn: normalized.range.endColumn,
    });
  }

  setColumnWidth(width: RenderColumnWidth): void {
    assertPositiveInteger(width.column, 'column width column');

    if (width.width < 0) {
      throw new ReportEngineError('Column width must be greater than 0.');
    }

    this.sheet.columnWidths.push({ ...width });
  }

  setColumnHidden(visibility: RenderColumnVisibility): void {
    assertPositiveInteger(visibility.column, 'column visibility column');
    this.sheet.columnVisibility.push({ ...visibility });
  }

  setRowHeight(height: RenderRowHeight): void {
    assertPositiveInteger(height.row, 'row height row');

    if (height.height < 0) {
      throw new ReportEngineError('Row height must be greater than 0.');
    }

    this.sheet.rowHeights.push({ ...height });
  }

  setRowHidden(visibility: RenderRowVisibility): void {
    assertPositiveInteger(visibility.row, 'row visibility row');
    this.sheet.rowVisibility.push({ ...visibility });
  }

  /** Trả về sheet đã được sắp xếp sẵn sàng cho ExcelJS renderer. */
  finish(): RenderPlanSheet {
    this.sheet.rows.sort((a, b) => a.index - b.index);

    for (const row of this.sheet.rows) {
      row.cells.sort((a, b) => a.column - b.column);
    }

    return this.sheet;
  }

  private getOrCreateRow(rowIndex: number): RenderPlanSheet['rows'][number] {
    const existing = this.rowIndex.get(rowIndex);

    if (existing !== undefined) {
      return this.sheet.rows[existing]!;
    }

    const row = { index: rowIndex, cells: [] };
    this.rowIndex.set(rowIndex, this.sheet.rows.length);
    this.sheet.rows.push(row);
    return row;
  }
}

function cloneCell(cell: RenderCell): RenderCell {
  return {
    ...cell,
    style: cell.style
      ? typeof cell.style === 'string'
        ? cell.style
        : (cloneStylePart(cell.style) as typeof cell.style)
      : undefined,
    inlineStyle: cell.inlineStyle ? (cloneStylePart(cell.inlineStyle) as typeof cell.inlineStyle) : undefined,
    formulaResult:
      cell.formulaResult !== undefined ? (cloneStylePart(cell.formulaResult) as typeof cell.formulaResult) : undefined,
    link: cell.link ? { ...cell.link } : undefined,
  };
}
