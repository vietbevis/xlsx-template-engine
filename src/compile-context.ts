import type { AddressRegistry } from './address-registry';
import type { SheetDefinition, WorkbookDefinition } from './types';
import type { VariableScope } from './variable-engine';

/**
 * Context nhỏ gọn truyền xuyên suốt quá trình compile một sheet.
 *
 * Thay thế `SheetContext` cũ (god object có cả workbook lẫn styles thừa).
 * `styles` được truy cập qua `workbook.styles` khi cần — không lưu thêm.
 */
export interface CompileContext {
  readonly workbook: WorkbookDefinition;
  readonly sheet: SheetDefinition;
  /** Số cột tối đa của sheet, dùng để resolve colSpan='remaining'. */
  readonly sheetColumnCount: number;
  readonly variables: VariableScope;
  /**
   * Registry tích lũy địa chỉ ô theo id trong single-pass.
   * Grid blocks ghi vào registry ngay khi layout được tính.
   * Table/section rows dùng để lookup cross-sheet refs.
   */
  readonly registry: AddressRegistry;
}
