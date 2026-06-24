import ExcelJS from 'exceljs';
import type { Writable } from 'stream';
import type { RenderCell, RenderPlan, RenderPlanSheet } from '../../compiler/render-plan';
import { ReportEngineError } from '../../core/errors';
import type { CellStyleDefinition, StyleValue } from '../../core/types';
import { cloneStylePart, isPlainObject } from '../../helpers/utils';

// ---------------------------------------------------------------------------
// Internal types
// ---------------------------------------------------------------------------

/** Minimal interface của worksheet chỉ cần khả năng lấy cell theo tọa độ,
 *  dùng để tái sử dụng `applyMergedStyleCoverage` cho cả streaming lẫn
 *  non-streaming workbook mà không cần ép kiểu toàn bộ ExcelJS.Worksheet. */
interface StyleableWorksheet {
  getCell(row: number, column: number): ExcelJS.Cell;
}

/** Index 2 chiều (rowIndex → columnIndex → RenderCell) để tra cứu O(1)
 *  thay vì duyệt mảng khi tìm master cell của một merge range. */
type CellPlanIndex = Map<number, Map<number, RenderCell>>;

/** Kiểu hợp lệ cho kết quả ước tính (cached result) của một formula cell. */
type FormulaResult = string | number | boolean | Date | ExcelJS.CellErrorValue;

/** Một merge range đã được gắn thêm style để áp dụng lên toàn bộ các ô bị
 *  che phủ bởi merge. ExcelJS chỉ lưu style tại master cell; muốn border
 *  hiển thị đúng trên các ô con ta phải ghi style thủ công. */
interface StyledMergeCoverage {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  style: CellStyleDefinition;
}

// ---------------------------------------------------------------------------
// Adapter
// ---------------------------------------------------------------------------

/**
 * Adapter chuyển đổi một `RenderPlan` (cấu trúc trung gian của report engine)
 * thành file Excel thực tế thông qua thư viện ExcelJS.
 *
 * Hỗ trợ ba phương thức xuất:
 * - `writeFile`   — ghi thẳng ra đường dẫn file (streaming, tiết kiệm RAM).
 * - `writeBuffer` — trả về `Buffer` trong bộ nhớ (phù hợp gửi qua HTTP).
 * - `writeStream` — pipe vào một `Writable` stream bất kỳ.
 *
 * Kiến trúc nội bộ:
 * - `createStreamingWorkbook` dùng `ExcelJS.stream.xlsx.WorkbookWriter` để ghi
 *   từng row rồi commit ngay, tránh giữ toàn bộ workbook trong RAM.
 * - `createWorkbook` dùng `ExcelJS.Workbook` thông thường (cần cho writeBuffer
 *   vì streaming writer không hỗ trợ `writeBuffer` trực tiếp).
 * - Cả hai path đều gọi chung các helper private để tránh trùng lặp logic.
 */
export class ExcelJsWorkbookAdapter {
  // -------------------------------------------------------------------------
  // Public API
  // -------------------------------------------------------------------------

  /** Ghi render plan ra file Excel tại `filePath` theo chế độ streaming. */
  async writeFile(renderPlan: RenderPlan, filePath: string): Promise<void> {
    const workbook = this.createStreamingWorkbook(renderPlan, { filename: filePath });
    await workbook.commit();
  }

  /** Render toàn bộ workbook vào RAM và trả về `Buffer`.
   *  Dùng khi cần gửi response HTTP hoặc upload lên object storage. */
  async writeBuffer(renderPlan: RenderPlan): Promise<Buffer> {
    const workbook = this.createWorkbook(renderPlan);
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
  }

  /** Pipe workbook vào `stream` theo chế độ streaming.
   *  Thích hợp khi muốn truyền trực tiếp xuống client mà không buffer. */
  async writeStream(renderPlan: RenderPlan, stream: Writable): Promise<void> {
    const workbook = this.createStreamingWorkbook(renderPlan, { stream });
    await workbook.commit();
  }

  // -------------------------------------------------------------------------
  // Workbook builders
  // -------------------------------------------------------------------------

  /**
   * Tạo và điền dữ liệu vào một streaming workbook (ExcelJS stream writer).
   *
   * Thứ tự xử lý mỗi sheet:
   * 1. Thiết lập độ rộng cột / ẩn cột.
   * 2. Thiết lập chiều cao hàng / ẩn hàng.
   * 3. Đăng ký các merge range.
   * 4. Ghi từng row, apply style cell, rồi commit row ngay lập tức.
   * 5. Apply border style lên các ô bị che bởi merge (xem `StyledMergeCoverage`).
   *
   * **Lưu ý streaming**: sau khi `row.commit()` được gọi, row đó bị flush ra
   * disk/stream và không thể chỉnh sửa lại. Vì vậy `applyMergedStyleCoverage`
   * phải được gọi TRƯỚC `row.commit()` — nhận `rowIndex` để chỉ xử lý đúng
   * hàng hiện tại, không phải toàn bộ merge coverage.
   */
  private createStreamingWorkbook(
    renderPlan: RenderPlan,
    options: Partial<ExcelJS.stream.xlsx.WorkbookStreamWriterOptions>,
  ): ExcelJS.stream.xlsx.WorkbookWriter {
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      ...options,
      useStyles: true,
    });

    if (renderPlan.metadata?.author) workbook.creator = renderPlan.metadata.author;

    for (const sheetPlan of renderPlan.sheets) {
      const sheet = workbook.addWorksheet(sheetPlan.name, {
        views: sheetPlan.views,
      });

      // Xây index cell plan và tính coverage trước để tái sử dụng khi ghi từng row.
      const cellPlanIndex = this.createCellPlanIndex(sheetPlan);
      const mergeCoverage = this.createStyledMergeCoverage(sheetPlan, renderPlan, cellPlanIndex);
      // Index coverage theo row để lookup O(1) thay vì quét toàn bộ mảng mỗi row.
      const mergeCoverageByRow = this.indexMergeCoverageByRow(mergeCoverage);

      for (const columnWidth of sheetPlan.columnWidths) {
        sheet.getColumn(columnWidth.column).width = columnWidth.width;
      }

      for (const columnVisibility of sheetPlan.columnVisibility ?? []) {
        sheet.getColumn(columnVisibility.column).hidden = columnVisibility.hidden;
      }

      for (const rowHeight of sheetPlan.rowHeights) {
        sheet.getRow(rowHeight.row).height = rowHeight.height;
      }

      for (const rowVisibility of sheetPlan.rowVisibility ?? []) {
        sheet.getRow(rowVisibility.row).hidden = rowVisibility.hidden;
      }

      for (const merge of sheetPlan.merges) {
        sheet.mergeCells(merge.startRow, merge.startColumn, merge.endRow, merge.endColumn);
      }

      for (const rowPlan of sheetPlan.rows) {
        const row = sheet.getRow(rowPlan.index);

        for (const cellPlan of rowPlan.cells) {
          this.applyCellPlan(row.getCell(cellPlan.column), cellPlan, renderPlan);
        }

        // Phải apply merge style TRƯỚC commit vì sau commit row không thể sửa.
        this.applyMergedStyleCoverage(
          sheet,
          mergeCoverageByRow.get(rowPlan.index) ?? [],
          rowPlan.index,
        );

        row.commit();
      }
    }

    this.applyDefinedNames(workbook, renderPlan);

    return workbook;
  }

  /**
   * Tạo và điền dữ liệu vào một workbook thông thường (in-memory).
   *
   * Khác với streaming, toàn bộ dữ liệu tồn tại trong RAM đến khi gọi
   * `writeBuffer`. Vì không cần commit từng row, `applyMergedStyleCoverage`
   * được gọi một lần cho toàn bộ sheet SAU khi tất cả rows đã được ghi.
   *
   * **Lưu ý thứ tự**: merge phải được đăng ký TRƯỚC khi ghi cell vào row,
   * vì ExcelJS có thể reset style của ô master khi merge được áp dụng sau.
   */
  private createWorkbook(renderPlan: RenderPlan): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();

    if (renderPlan.metadata?.author) workbook.creator = renderPlan.metadata.author;

    for (const sheetPlan of renderPlan.sheets) {
      const sheet = workbook.addWorksheet(sheetPlan.name, {
        views: sheetPlan.views,
      });

      const cellPlanIndex = this.createCellPlanIndex(sheetPlan);
      const mergeCoverage = this.createStyledMergeCoverage(sheetPlan, renderPlan, cellPlanIndex);

      for (const columnWidth of sheetPlan.columnWidths) {
        sheet.getColumn(columnWidth.column).width = columnWidth.width;
      }

      for (const columnVisibility of sheetPlan.columnVisibility ?? []) {
        sheet.getColumn(columnVisibility.column).hidden = columnVisibility.hidden;
      }

      for (const rowHeight of sheetPlan.rowHeights) {
        sheet.getRow(rowHeight.row).height = rowHeight.height;
      }

      for (const rowVisibility of sheetPlan.rowVisibility ?? []) {
        sheet.getRow(rowVisibility.row).hidden = rowVisibility.hidden;
      }

      for (const rowPlan of sheetPlan.rows) {
        const row = sheet.getRow(rowPlan.index);

        for (const cellPlan of rowPlan.cells) {
          this.applyCellPlan(row.getCell(cellPlan.column), cellPlan, renderPlan);
        }
      }

      for (const merge of sheetPlan.merges) {
        sheet.mergeCells(merge.startRow, merge.startColumn, merge.endRow, merge.endColumn);
      }

      // Với non-streaming workbook, apply toàn bộ coverage sau khi sheet đã đầy đủ.
      this.applyMergedStyleCoverage(sheet, mergeCoverage);
    }

    this.applyDefinedNames(workbook, renderPlan);

    return workbook;
  }

  // -------------------------------------------------------------------------
  // Cell & style helpers
  // -------------------------------------------------------------------------

  /**
   * Áp dụng toàn bộ thông tin từ `cellPlan` lên một `ExcelJS.Cell`:
   * - Nếu có formula, tạo `CellFormulaValue` kèm cached result (nếu có)
   *   để Excel không cần tính lại khi mở file.
   * - Resolve style theo thứ tự ưu tiên: defaultStyle → namedStyle → inlineStyle.
   */
  private applyCellPlan(
    cell: ExcelJS.Cell,
    cellPlan: RenderPlan['sheets'][number]['rows'][number]['cells'][number],
    renderPlan: RenderPlan,
  ): void {
    cell.value = cellPlan.formula
      ? this.createFormulaValue(cellPlan.formula, cellPlan.formulaResult ?? cellPlan.value)
      : (cellPlan.value ?? null);

    const style = this.resolveCellStyle(cellPlan, renderPlan);

    if (style) {
      cell.style = cloneStylePart(style) as Partial<ExcelJS.Style>;
    }
  }

  /**
   * Áp dụng style của master cell lên tất cả ô bị merge che phủ.
   *
   * ExcelJS chỉ lưu style tại ô master (top-left) của một merge range.
   * Khi mở file, Excel hiển thị border dựa trên style của từng ô riêng lẻ,
   * không kế thừa từ master. Do đó ta phải ghi style thủ công lên toàn bộ
   * ô trong vùng merge.
   *
   * @param sheet
   * @param mergeCoverage
   * @param rowIndex - Nếu được cung cấp, chỉ xử lý các ô thuộc hàng này
   *   (dùng trong streaming mode để không vi phạm giới hạn commit-per-row).
   *   Nếu `undefined`, xử lý toàn bộ vùng merge (dùng trong non-streaming mode).
   */
  private applyMergedStyleCoverage(
    sheet: StyleableWorksheet,
    mergeCoverage: readonly StyledMergeCoverage[],
    rowIndex?: number,
  ): void {
    for (const merge of mergeCoverage) {
      const startRow = rowIndex ?? merge.startRow;
      const endRow = rowIndex ?? merge.endRow;

      for (let row = startRow; row <= endRow; row += 1) {
        for (let column = merge.startColumn; column <= merge.endColumn; column += 1) {
          sheet.getCell(row, column).style = cloneStylePart(merge.style) as Partial<ExcelJS.Style>;
        }
      }
    }
  }

  /**
   * Xây dựng index 2 chiều `rowIndex → columnIndex → RenderCell` từ sheet plan.
   * Cho phép tra cứu master cell của bất kỳ merge range nào trong O(1).
   */
  private createCellPlanIndex(sheetPlan: RenderPlanSheet): CellPlanIndex {
    const index: CellPlanIndex = new Map();

    for (const row of sheetPlan.rows) {
      const cellsByColumn = new Map<number, RenderCell>();

      for (const cell of row.cells) {
        cellsByColumn.set(cell.column, cell);
      }

      index.set(row.index, cellsByColumn);
    }

    return index;
  }

  /**
   * Lọc ra các merge range cần áp dụng border coverage và gắn style vào.
   *
   * Chỉ những merge range có border style mới cần xử lý đặc biệt — vì đây là
   * trường hợp duy nhất mà việc ExcelJS chỉ lưu style tại master cell tạo ra
   * kết quả hiển thị sai trong Excel (border bị thiếu ở các ô con).
   */
  private createStyledMergeCoverage(
    sheetPlan: RenderPlanSheet,
    renderPlan: RenderPlan,
    cellPlanIndex: CellPlanIndex,
  ): StyledMergeCoverage[] {
    const coverage: StyledMergeCoverage[] = [];

    for (const merge of sheetPlan.merges) {
      // Chỉ master cell (top-left) của merge mới mang style đại diện.
      const masterCellPlan = cellPlanIndex.get(merge.startRow)?.get(merge.startColumn);

      if (!masterCellPlan?.style && !masterCellPlan?.inlineStyle) {
        continue;
      }

      const style = this.resolveCellStyle(masterCellPlan, renderPlan);

      // Bỏ qua merge range không có border — không cần clone style xuống ô con.
      if (!style?.border) {
        continue;
      }

      coverage.push({ ...merge, style });
    }

    return coverage;
  }

  /**
   * Tạo index `rowIndex → StyledMergeCoverage[]` để streaming writer có thể
   * tra cứu nhanh tất cả merge coverage cần apply khi commit một row cụ thể.
   */
  private indexMergeCoverageByRow(
    mergeCoverage: readonly StyledMergeCoverage[],
  ): Map<number, StyledMergeCoverage[]> {
    const index = new Map<number, StyledMergeCoverage[]>();

    for (const merge of mergeCoverage) {
      for (let row = merge.startRow; row <= merge.endRow; row += 1) {
        const rowMerges = index.get(row) ?? [];
        rowMerges.push(merge);
        index.set(row, rowMerges);
      }
    }

    return index;
  }

  /**
   * Tạo `CellFormulaValue` cho ExcelJS từ công thức và kết quả ước tính.
   *
   * Khi `result` có giá trị, Excel sẽ hiển thị cached result ngay khi mở file
   * mà không cần tính toán lại — quan trọng với các file được mở trên môi trường
   * không có calculation engine đầy đủ (ví dụ: Excel Online, LibreOffice).
   */
  private createFormulaValue(formula: string, result: unknown): ExcelJS.CellValue {
    if (result !== undefined && result !== null) {
      return {
        formula,
        result: result as FormulaResult,
        date1904: false,
      };
    }

    return { formula, date1904: false };
  }

  /**
   * Đăng ký tất cả named ranges từ render plan vào workbook dưới dạng
   * Excel Defined Names với địa chỉ tuyệt đối (dấu `$`), ví dụ:
   * `'Sheet1'!$A$1:$C$5`.
   *
   * Named ranges cho phép template, macro, hoặc công thức tham chiếu đến
   * vùng dữ liệu mà không phụ thuộc vào địa chỉ ô cụ thể.
   */
  private applyDefinedNames(
    workbook: ExcelJS.Workbook | ExcelJS.stream.xlsx.WorkbookWriter,
    renderPlan: RenderPlan,
  ): void {
    for (const namedRange of renderPlan.namedRanges ?? []) {
      workbook.definedNames.add(
        `${quoteSheetName(namedRange.sheetName)}!${formatAbsoluteCell(namedRange.startRow, namedRange.startColumn)}:${formatAbsoluteCell(namedRange.endRow, namedRange.endColumn)}`,
        namedRange.name,
      );
    }
  }

  // -------------------------------------------------------------------------
  // Style resolution
  // -------------------------------------------------------------------------

  /**
   * Resolve style cuối cùng của một cell theo thứ tự ưu tiên tăng dần:
   *   1. `renderPlan.defaultStyle`  — style mặc định toàn workbook.
   *   2. `cellPlan.style`           — named style hoặc style object từ render plan.
   *   3. `cellPlan.inlineStyle`     — style inline ghi đè cục bộ tại cell.
   *
   * Các lớp style được deep-merge: key nào xuất hiện ở lớp sau sẽ override
   * lớp trước; các object lồng nhau được merge đệ quy thay vì replace hoàn toàn.
   */
  private resolveCellStyle(
    cellPlan: RenderCell,
    renderPlan: RenderPlan,
  ): CellStyleDefinition | undefined {
    const baseStyle = this.resolveStyleValue(cellPlan.style, renderPlan);
    const style = this.mergeCellStyles(renderPlan.defaultStyle, baseStyle);

    return this.mergeCellStyles(style, cellPlan.inlineStyle);
  }

  /**
   * Resolve một `StyleValue` thành `CellStyleDefinition`:
   * - Nếu là `string` → tra cứu trong `renderPlan.styles` registry; ném lỗi
   *   nếu không tìm thấy để phát hiện typo sớm.
   * - Nếu là object → trả về trực tiếp.
   * - Nếu `undefined` → trả về `undefined`.
   */
  private resolveStyleValue(
    style: StyleValue | undefined,
    renderPlan: RenderPlan,
  ): CellStyleDefinition | undefined {
    if (!style) {
      return undefined;
    }

    if (typeof style !== 'string') {
      return style;
    }

    const registryStyle = renderPlan.styles?.[style];

    if (!registryStyle) {
      throw new ReportEngineError(`Render plan references unknown style "${style}".`);
    }

    return registryStyle;
  }

  /**
   * Deep-merge `override` lên trên `base`, trả về style mới.
   * Nếu cả hai đều `undefined`, trả về `undefined`.
   * Nếu chỉ một bên có giá trị, trả về `{}` (empty style object) để đảm bảo
   * kiểu trả về luôn là object khi ít nhất một bên không null.
   */
  private mergeCellStyles(
    base: CellStyleDefinition | undefined,
    override: CellStyleDefinition | undefined,
  ): CellStyleDefinition | undefined {
    if (!base && !override) {
      return undefined;
    }

    return this.mergeStylePart(base, override) ?? {};
  }

  /**
   * Deep-merge generic `override` lên `base`, clone các giá trị primitive/array
   * từ override để tránh shared reference giữa các cell.
   *
   * Quy tắc merge:
   * - Nếu cả `base[key]` và `override[key]` đều là plain object → merge đệ quy.
   * - Ngược lại → clone và ghi đè hoàn toàn bằng giá trị của override.
   */
  private mergeStylePart<T extends Record<string, unknown>>(
    base: T | undefined,
    override: T | undefined,
  ): T | undefined {
    if (!base && !override) {
      return undefined;
    }

    const merged: Record<string, unknown> = { ...(base ?? {}) };

    for (const [key, value] of Object.entries(override ?? {})) {
      const baseValue = merged[key];

      if (isPlainObject(baseValue) && isPlainObject(value)) {
        merged[key] = this.mergeStylePart(baseValue, value);
        continue;
      }

      merged[key] = cloneStylePart(value);
    }

    return merged as T;
  }
}

// ---------------------------------------------------------------------------
// Module-level helpers (pure functions, không phụ thuộc state của adapter)
// ---------------------------------------------------------------------------

/**
 * Định dạng địa chỉ ô tuyệt đối theo ký hiệu A1, ví dụ: `$C$5`.
 * Dùng khi tạo địa chỉ cho defined names.
 */
function formatAbsoluteCell(row: number, column: number): string {
  return `$${columnNumberToName(column)}$${row}`;
}

/**
 * Chuyển số cột (1-based) thành tên cột Excel theo hệ cơ số 26 biến thể,
 * ví dụ: 1 → "A", 26 → "Z", 27 → "AA", 703 → "AAA".
 *
 * Khác hệ base-26 thông thường ở chỗ không có ký tự "0" — 'Z' (26) là
 * chữ số cuối của mỗi vị trí, nên phép chia dùng `(remaining - 1) % 26`.
 */
function columnNumberToName(column: number): string {
  let remaining = column;
  let name = '';

  while (remaining > 0) {
    const modulo = (remaining - 1) % 26;
    name = String.fromCharCode(65 + modulo) + name;
    remaining = Math.floor((remaining - modulo) / 26);
  }

  return name;
}

/**
 * Bọc tên sheet trong dấu nháy đơn theo cú pháp Excel, escape dấu `'` nội tại
 * bằng cách nhân đôi (`''`), ví dụ: `O'Brien` → `'O''Brien'`.
 * Cần thiết khi tên sheet xuất hiện trong địa chỉ công thức hoặc defined names.
 */
function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replace(/'/g, "''")} '`;
}
