import { ReportEngineError } from '../core/errors';
import type {
  CellStyleDefinition,
  StyleRegistry,
  StyleValue,
  WorkbookMetadata,
} from '../core/types';
import type {
  RenderCell,
  RenderColumnVisibility,
  RenderColumnWidth,
  RenderMergeRange,
  RenderPlan,
  RenderPlanSheet,
  RenderRowHeight,
  RenderRowVisibility,
  RenderSheetView,
  ResolvedNamedRange,
} from './render-plan';
import { assertMergeDoesNotOverlap, type MergeRange, normalizeMergeRange } from './merge-engine';
import { cloneStylePart } from '../helpers/utils';

// ---------------------------------------------------------------------------
// Builder options
// ---------------------------------------------------------------------------

/** Các tùy chọn cấp workbook được truyền vào khi khởi tạo `RenderPlanBuilder`.
 *  Tất cả đều optional — builder hoạt động bình thường nếu không có tùy chọn nào. */
export interface RenderPlanBuilderOptions {
  /** Metadata của workbook (tác giả, tiêu đề, từ khóa, ...). */
  metadata?: WorkbookMetadata;
  /** Style mặc định áp dụng cho mọi cell trừ khi bị override. */
  defaultStyle?: CellStyleDefinition;
  /** Registry các named style dùng chung, tham chiếu bằng tên string. */
  styles?: StyleRegistry;
  /** Danh sách named ranges đã được resolve thành tọa độ tuyệt đối. */
  namedRanges?: ResolvedNamedRange[];
}

// ---------------------------------------------------------------------------
// Builder
// ---------------------------------------------------------------------------

/**
 * Builder tích lũy dữ liệu từng bước rồi tạo ra một `RenderPlan` bất biến.
 *
 * **Luồng sử dụng điển hình:**
 * ```
 * const builder = new RenderPlanBuilder({ defaultStyle, styles });
 * builder.addSheet('sheet1', 'Danh sách');
 * builder.addCell('sheet1', { row: 1, column: 1, value: 'Hello' });
 * builder.addMerge('sheet1', { startRow: 1, startColumn: 1, endRow: 1, endColumn: 3 });
 * const plan = builder.build();
 * ```
 *
 * **Thiết kế nội bộ:**
 * - `sheets` — Map từ sheetId → `RenderPlanSheet`, giữ thứ tự thêm vào (insertion order).
 * - `rowsBySheet` — Index phụ `sheetId → rowIndex → row` để `addCell` tra cứu/tạo row trong O(1)
 *   thay vì quét mảng `sheet.rows` mỗi lần.
 * - `mergeRanges` — Danh sách phẳng tất cả merge ranges của mọi sheet, dùng để phát hiện
 *   chồng lấp xuyên-sheet một cách hiệu quả.
 *
 * **Tính bất biến của output:** `build()` deep-clone toàn bộ dữ liệu trước khi trả về,
 * đảm bảo caller không thể vô tình mutate trạng thái nội bộ của builder.
 */
export class RenderPlanBuilder {
  private readonly sheets = new Map<string, RenderPlanSheet>();
  /** Index phụ để getOrCreateRow chạy O(1), tránh `.find()` trên mảng rows mỗi lần addCell. */
  private readonly rowsBySheet = new Map<string, Map<number, RenderPlanSheet['rows'][number]>>();
  /** Danh sách tất cả MergeRange đã đăng ký (mọi sheet) để kiểm tra chồng lấp toàn cục. */
  private readonly mergeRanges: MergeRange[] = [];

  constructor(private readonly options: RenderPlanBuilderOptions = {}) {}

  // -------------------------------------------------------------------------
  // Sheet
  // -------------------------------------------------------------------------

  /**
   * Thêm một sheet mới vào render plan.
   *
   * @param id   - Định danh nội bộ duy nhất của sheet, dùng khi gọi các method khác.
   * @param name - Tên hiển thị trên tab Excel.
   * @param options
   * @param options.views - Cấu hình view (freeze panes, zoom, ...).
   * @throws Nếu `id` đã tồn tại trong builder.
   */
  addSheet(id: string, name: string, options: { views?: RenderSheetView[] } = {}): void {
    if (this.sheets.has(id)) {
      throw new ReportEngineError(`Render plan already contains sheet "${id}".`);
    }

    this.sheets.set(id, {
      id,
      name,
      views: options.views ? options.views.map((view) => ({ ...view })) : undefined,
      rows: [],
      merges: [],
      columnWidths: [],
      columnVisibility: [],
      rowHeights: [],
      rowVisibility: [],
    });
    this.rowsBySheet.set(id, new Map());
  }

  // -------------------------------------------------------------------------
  // Cell
  // -------------------------------------------------------------------------

  /**
   * Thêm một cell vào sheet. Nếu row chưa tồn tại sẽ được tạo tự động.
   *
   * Cell được deep-clone ngay khi thêm vào để tránh caller vô tình mutate
   * dữ liệu sau khi đã đẩy vào builder.
   *
   * @throws Nếu `cell.row` hoặc `cell.column` không phải số nguyên dương.
   * @throws Nếu `sheetId` không tồn tại.
   */
  addCell(sheetId: string, cell: RenderCell): void {
    assertPositiveInteger(cell.row, 'cell row');
    assertPositiveInteger(cell.column, 'cell column');

    const sheet = this.getSheet(sheetId);
    const row = this.getOrCreateRow(sheet, cell.row);
    row.cells.push({
      ...cell,
      style: cloneStyleValue(cell.style),
      inlineStyle: cell.inlineStyle ? cloneStyle(cell.inlineStyle) : undefined,
      formulaResult: cloneCellValue(cell.formulaResult),
      link: cell.link ? { ...cell.link } : undefined,
    });
  }

  // -------------------------------------------------------------------------
  // Merge
  // -------------------------------------------------------------------------

  /**
   * Đăng ký một merge range cho sheet.
   *
   * Quy trình xử lý:
   * 1. Normalize range — validate tọa độ và bỏ qua nếu là ô đơn.
   * 2. Kiểm tra không chồng lấp với bất kỳ merge nào đã đăng ký (kể cả sheet khác).
   * 3. Lưu vào `mergeRanges` (danh sách toàn cục) và `sheet.merges` (danh sách cục bộ).
   *
   * @throws Nếu range có tọa độ không hợp lệ hoặc chồng lấp với merge đã tồn tại.
   */
  addMerge(sheetId: string, range: RenderMergeRange): void {
    const sheet = this.getSheet(sheetId);
    const normalized = normalizeMergeRange(sheetId, range);

    if (normalized.type === 'skip-single-cell') {
      return;
    }

    const normalizedRange = normalized.range;
    assertMergeDoesNotOverlap(normalizedRange, this.mergeRanges);
    this.mergeRanges.push(normalizedRange);
    sheet.merges.push({
      startRow: normalizedRange.startRow,
      startColumn: normalizedRange.startColumn,
      endRow: normalizedRange.endRow,
      endColumn: normalizedRange.endColumn,
    });
  }

  // -------------------------------------------------------------------------
  // Column / Row metadata
  // -------------------------------------------------------------------------

  /**
   * Đặt độ rộng cho một cột.
   * @throws Nếu `width.column` không phải số nguyên dương hoặc `width.width` âm.
   */
  setColumnWidth(sheetId: string, width: RenderColumnWidth): void {
    assertPositiveInteger(width.column, 'column width column');

    if (width.width < 0) {
      throw new ReportEngineError('Column width must be greater 0.');
    }

    this.getSheet(sheetId).columnWidths.push({ ...width });
  }

  /**
   * Ẩn hoặc hiện một cột.
   * @throws Nếu `visibility.column` không phải số nguyên dương.
   */
  setColumnHidden(sheetId: string, visibility: RenderColumnVisibility): void {
    assertPositiveInteger(visibility.column, 'column visibility column');
    this.getSheet(sheetId).columnVisibility.push({ ...visibility });
  }

  /**
   * Đặt chiều cao cho một hàng.
   * @throws Nếu `height.row` không phải số nguyên dương hoặc `height.height` âm.
   */
  setRowHeight(sheetId: string, height: RenderRowHeight): void {
    assertPositiveInteger(height.row, 'row height row');

    if (height.height < 0) {
      throw new ReportEngineError('Row height must be greater 0.');
    }

    this.getSheet(sheetId).rowHeights.push({ ...height });
  }

  /**
   * Ẩn hoặc hiện một hàng.
   * @throws Nếu `visibility.row` không phải số nguyên dương.
   */
  setRowHidden(sheetId: string, visibility: RenderRowVisibility): void {
    assertPositiveInteger(visibility.row, 'row visibility row');
    this.getSheet(sheetId).rowVisibility.push({ ...visibility });
  }

  // -------------------------------------------------------------------------
  // Build
  // -------------------------------------------------------------------------

  /**
   * Hoàn tất quá trình build và trả về `RenderPlan` bất biến.
   *
   * Toàn bộ dữ liệu được deep-clone một lần nữa để đảm bảo output hoàn toàn
   * độc lập với trạng thái nội bộ của builder — caller có thể giữ và truyền
   * `RenderPlan` đi mà không lo bị ảnh hưởng nếu builder tiếp tục bị gọi.
   *
   * Rows và cells được sort theo index/column tăng dần trước khi đưa vào plan,
   * đảm bảo adapter nhận dữ liệu theo thứ tự ổn định bất kể thứ tự `addCell`
   * được gọi.
   */
  build(): RenderPlan {
    return {
      metadata: this.options.metadata
        ? {
            ...this.options.metadata,
            keywords: this.options.metadata.keywords
              ? [...this.options.metadata.keywords]
              : undefined,
          }
        : undefined,
      defaultStyle: this.options.defaultStyle ? cloneStyle(this.options.defaultStyle) : undefined,
      styles: this.options.styles ? resolveStyleRegistry(this.options.styles) : undefined,
      namedRanges: this.options.namedRanges
        ? this.options.namedRanges.map((range) => ({ ...range }))
        : undefined,
      sheets: Array.from(this.sheets.values()).map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        views: sheet.views ? sheet.views.map((view) => ({ ...view })) : undefined,
        // Sort rows theo row index và cells theo column index tăng dần.
        rows: [...sheet.rows].sort(compareRows).map((row) => ({
          index: row.index,
          cells: [...row.cells].sort(compareCells).map((cell) => ({
            ...cell,
            style: cloneStyleValue(cell.style),
            inlineStyle: cell.inlineStyle ? cloneStyle(cell.inlineStyle) : undefined,
            formulaResult: cloneCellValue(cell.formulaResult),
            link: cell.link ? { ...cell.link } : undefined,
          })),
        })),
        merges: sheet.merges.map((merge) => ({ ...merge })),
        columnWidths: sheet.columnWidths.map((width) => ({ ...width })),
        columnVisibility: sheet.columnVisibility.map((visibility) => ({ ...visibility })),
        rowHeights: sheet.rowHeights.map((height) => ({ ...height })),
        rowVisibility: sheet.rowVisibility.map((visibility) => ({ ...visibility })),
      })),
    };
  }

  // -------------------------------------------------------------------------
  // Private helpers
  // -------------------------------------------------------------------------

  /** Lấy sheet theo `sheetId`, ném lỗi nếu không tìm thấy. */
  private getSheet(sheetId: string): RenderPlanSheet {
    const sheet = this.sheets.get(sheetId);

    if (!sheet) {
      throw new ReportEngineError(`Render plan does not contain sheet "${sheetId}".`);
    }

    return sheet;
  }

  /**
   * Lấy row tại `rowIndex` trong sheet; nếu chưa tồn tại thì tạo mới và
   * đồng thời đăng ký vào cả `sheet.rows` lẫn index phụ `rowsBySheet`.
   *
   * Dùng index phụ `rowsBySheet` để tra cứu O(1) — tránh `.find()` O(n)
   * trên `sheet.rows` khi số lượng row lớn (báo cáo hàng nghìn dòng).
   *
   * @throws Nếu `rowsBySheet` không chứa entry cho sheet (không nên xảy ra
   *   nếu `addSheet` và `getSheet` luôn được gọi đúng thứ tự).
   */
  private getOrCreateRow(
    sheet: RenderPlanSheet,
    rowIndex: number,
  ): RenderPlanSheet['rows'][number] {
    const rowMap = this.rowsBySheet.get(sheet.id);

    if (!rowMap) {
      throw new ReportEngineError(`Render plan does not contain sheet "${sheet.id}".`);
    }

    const existingRow = rowMap.get(rowIndex);

    if (existingRow) {
      return existingRow;
    }

    const row = { index: rowIndex, cells: [] };
    sheet.rows.push(row);
    rowMap.set(rowIndex, row);
    return row;
  }
}

// ---------------------------------------------------------------------------
// Module-level helpers
// ---------------------------------------------------------------------------

/** So sánh hai row theo index tăng dần — dùng để sort trong `build()`. */
function compareRows(
  left: RenderPlanSheet['rows'][number],
  right: RenderPlanSheet['rows'][number],
): number {
  return left.index - right.index;
}

/** So sánh hai cell theo column tăng dần — dùng để sort trong `build()`. */
function compareCells(left: RenderCell, right: RenderCell): number {
  return left.column - right.column;
}

/** Assert một giá trị phải là số nguyên dương (>= 1).
 *  Tái sử dụng cho mọi tọa độ row/column trước khi đưa vào plan. */
function assertPositiveInteger(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new ReportEngineError(`Render plan ${label} must be a positive integer.`);
  }
}

/**
 * Deep-clone toàn bộ style registry, đảm bảo các style object trong plan
 * không share reference với object được truyền vào từ ngoài.
 */
function resolveStyleRegistry(styles: StyleRegistry): StyleRegistry {
  return Object.fromEntries(
    Object.entries(styles).map(([styleName, style]) => [styleName, cloneStyle(style)]),
  );
}

/**
 * Clone một `StyleValue`:
 * - String (named style reference) → trả về nguyên (immutable).
 * - Object (`CellStyleDefinition`) → deep-clone.
 * - `undefined` → trả về `undefined`.
 */
function cloneStyleValue(style: StyleValue | undefined): StyleValue | undefined {
  if (typeof style === 'string' || style === undefined) {
    return style;
  }

  return cloneStyle(style);
}

/** Deep-clone một `CellStyleDefinition` thông qua `cloneStylePart` generic. */
function cloneStyle(style: CellStyleDefinition): CellStyleDefinition {
  return cloneStylePart(style) as CellStyleDefinition;
}

/** Clone giá trị cached formula result (có thể là Date, number, string, object).
 *  Dùng `cloneStylePart` vì cùng cơ chế deep-clone plain object / array. */
function cloneCellValue(value: RenderCell['formulaResult']): RenderCell['formulaResult'] {
  return cloneStylePart(value) as RenderCell['formulaResult'];
}
