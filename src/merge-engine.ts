import { ReportEngineError } from './errors';
import { assertPositiveInteger } from './helpers/common';
import type { RenderMergeRange } from './render-plan';

export interface MergeRange extends RenderMergeRange {
  sheetId: string;
}

export type NormalizedMergeRange = { type: 'skip-single-cell' } | { type: 'range'; range: MergeRange };

/**
 * Chuẩn hóa một merge range từ render plan thành dạng có thể sử dụng để ghi vào sheet.
 *
 * - Validate các tọa độ phải là số nguyên dương và end >= start.
 * - Nếu range chỉ gồm đúng 1 ô (startRow === endRow và startColumn === endColumn),
 *   trả về `skip-single-cell` để bỏ qua (không cần merge ô đơn).
 * - Ngược lại, gắn thêm `sheetId` và trả về range hợp lệ.
 */
export function normalizeMergeRange(sheetId: string, range: RenderMergeRange): NormalizedMergeRange {
  assertPositiveInteger(range.startRow, 'Render plan merge start row');
  assertPositiveInteger(range.startColumn, 'Render plan merge start column');
  assertPositiveInteger(range.endRow, 'Render plan merge end row');
  assertPositiveInteger(range.endColumn, 'Render plan merge end column');

  if (range.endRow < range.startRow || range.endColumn < range.startColumn) {
    throw new ReportEngineError('Merge range end must be greater than or equal to start.');
  }

  if (range.startRow === range.endRow && range.startColumn === range.endColumn) {
    return { type: 'skip-single-cell' };
  }

  return {
    type: 'range',
    range: {
      sheetId,
      startRow: range.startRow,
      startColumn: range.startColumn,
      endRow: range.endRow,
      endColumn: range.endColumn,
    },
  };
}

/**
 * Kiểm tra một merge range mới (candidate) có chồng lấp với bất kỳ range nào
 * trong danh sách đã tồn tại hay không.
 *
 * Chỉ so sánh các range cùng `sheetId`. Nếu phát hiện chồng lấp, ném lỗi
 * kèm thông tin tọa độ của cả hai range để dễ debug.
 */
export function assertMergeDoesNotOverlap(candidate: MergeRange, existingRanges: MergeRange[]): void {
  const overlappingRange = existingRanges.find(
    (existing) => existing.sheetId === candidate.sheetId && rangesOverlap(existing, candidate),
  );

  if (overlappingRange) {
    throw new ReportEngineError(
      `Merge range ${formatMergeRange(candidate)} overlaps existing range ${formatMergeRange(overlappingRange)}.`,
    );
  }
}

/**
 * Kiểm tra hai RenderMergeRange có chồng lấp nhau không (AABB intersection).
 *
 * Hai range KHÔNG chồng lấp khi một trong các điều kiện sau thỏa mãn:
 *   - left hoàn toàn ở trên right  (left.endRow < right.startRow)
 *   - right hoàn toàn ở trên left  (right.endRow < left.startRow)
 *   - left hoàn toàn ở trái right  (left.endColumn < right.startColumn)
 *   - right hoàn toàn ở trái left  (right.endColumn < left.startColumn)
 * Phủ định của tất cả điều kiện trên → hai range chồng lấp.
 */
function rangesOverlap(left: RenderMergeRange, right: RenderMergeRange): boolean {
  return !(
    left.endRow < right.startRow ||
    right.endRow < left.startRow ||
    left.endColumn < right.startColumn ||
    right.endColumn < left.startColumn
  );
}

/**
 * Định dạng một MergeRange thành chuỗi dễ đọc theo ký hiệu R1C1 của Excel, R = row, C = column
 * ví dụ: `"Sheet1"!R2C1:R4C3`.
 * Dùng cho mục đích hiển thị trong thông báo lỗi.
 */
function formatMergeRange(range: MergeRange): string {
  return `"${range.sheetId}"!R${range.startRow}C${range.startColumn}:R${range.endRow}C${range.endColumn}`;
}
