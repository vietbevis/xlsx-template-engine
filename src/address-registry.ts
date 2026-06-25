import { ReportEngineError } from './errors';
import { createFormulaId, type CellAddress } from './formula-engine';

/**
 * Mutable registry tích lũy địa chỉ ô (row, column) theo id trong suốt
 * quá trình compile single-pass.
 *
 * Thay thế hoàn toàn pre-pass `collectWorkbookFormulaIds` trong compile.ts cũ.
 * Mỗi lần `compileBlock` ghi ô có `id`, nó gọi `register()` ngay tại chỗ.
 * Lookup sau đó qua `resolve()`.
 */
export class AddressRegistry {
  private readonly map = new Map<string, CellAddress>();

  register(sheetId: string, sheetName: string, id: string, row: number, column: number): void {
    const key = createFormulaId(sheetId, id);

    if (this.map.has(key)) {
      throw new ReportEngineError(`Duplicate formula cell id "${id}" in sheet "${sheetId}".`);
    }

    this.map.set(key, { row, column, sheetId, sheetName });
  }

  resolve(id: string, sheetId?: string): CellAddress | undefined {
    return this.map.get(createFormulaId(sheetId, id));
  }

  /** Trả về snapshot bất biến để truyền vào FormulaCompileContext. */
  snapshot(): ReadonlyMap<string, CellAddress> {
    return this.map;
  }
}
