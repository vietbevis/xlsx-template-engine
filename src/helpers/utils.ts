import { ReportEngineError } from '../core/errors';
import { TableColumnNode, TableLeafColumn } from '../core/types';

/**
 * Deep-clone một phần của style object để tránh các cell dùng chung reference.
 * - Array → clone đệ quy từng phần tử.
 * - Plain object → clone đệ quy từng entry.
 * - Primitive / Date / ... → trả về nguyên giá trị (immutable).
 */
export function cloneStylePart(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((item) => cloneStylePart(item));
  }

  if (!isPlainObject(value)) {
    return value;
  }

  return Object.fromEntries(Object.entries(value).map(([key, childValue]) => [key, cloneStylePart(childValue)]));
}

/**
 * Kiểm tra `value` có phải plain object không (loại trừ null và array).
 * Dùng làm type guard để phân biệt object cần merge đệ quy với giá trị scalar.
 */
export function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/**
 * Kiểm tra `value` có phải là số nguyên >= 1 không, nếu không thì throw lỗi
 * @param value
 * @param label
 */
export function assertPositiveInteger(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new ReportEngineError(`Layout cursor ${label} must be a positive integer.`);
  }
}

/**
 * Làm phẳng cây cột (có thể lồng nhau) thành danh sách leaf column.
 * Chỉ leaf column (không có children) mới tương ứng với cột thực trong Excel.
 */
export function flattenColumns(columns: readonly TableColumnNode[]): TableLeafColumn[] {
  return columns.flatMap((column) =>
    column.children?.length ? flattenColumns(column.children) : [column as TableLeafColumn],
  );
}

/**
 * Tính độ sâu của header (số hàng header cần render).
 * Ví dụ: nếu có group column lồng 2 cấp → headerDepth = 2.
 */
export function calculateTableHeaderDepth(columns: readonly TableColumnNode[]): number {
  return Math.max(
    ...columns.map((column) => (column.children?.length ? 1 + calculateTableHeaderDepth(column.children) : 1)),
  );
}
