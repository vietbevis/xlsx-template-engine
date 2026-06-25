import { CompileError } from '../errors';
import type { CellAddress } from './formula-compiler';

/**
 * Registry tích lũy địa chỉ ô (row, column) theo id trong suốt quá trình
 * compile single-pass. Mỗi lần `compileBlock` ghi ô có `id`, nó gọi
 * `register()` ngay tại chỗ. Lookup sau đó qua `resolve()`.
 */
export class AddressRegistry {
  private readonly map = new Map<string, CellAddress>();

  register(sheetId: string, sheetName: string, id: string, row: number, column: number): void {
    const key = `${sheetId}:${id}`;

    if (this.map.has(key)) {
      throw new CompileError(`Duplicate formula cell id "${id}" in sheet "${sheetId}".`, { sheetId, id });
    }

    this.map.set(key, { row, column, sheetId, sheetName });
  }

  resolve(id: string, sheetId?: string): CellAddress | undefined {
    return this.map.get(sheetId ? `${sheetId}:${id}` : id);
  }
}
