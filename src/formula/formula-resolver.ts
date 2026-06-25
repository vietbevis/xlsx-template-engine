import type { AddressRegistry } from './address-registry';
import { FormulaError } from '../errors';
import { FormulaCompiler, type CellAddress, type FormulaCompileContext } from './formula-compiler';
import type { FormulaRangeScope } from '../types';

// ─── Public interfaces ────────────────────────────────────────────────────────

export interface RenderedDataRow {
  data: Record<string, unknown>;
  row: number;
}

export interface TableFormulaContextOptions {
  columnIdMap: Map<string, number>;
  row: number;
  firstColumn: number;
  registry: AddressRegistry;
  currentSheetId: string;
  scopedRows?: Partial<Record<FormulaRangeScope, RenderedDataRow[]>>;
}

// ─── FormulaResolver ──────────────────────────────────────────────────────────

export class FormulaResolver {
  static createGridContext(registry: AddressRegistry, currentSheetId: string): FormulaCompileContext {
    return {
      resolveCellId(id, sheetId) {
        return FormulaResolver.resolveRegisteredCell(registry, id, sheetId ?? currentSheetId, currentSheetId);
      },
      resolveRangeIds(startId, endId, sheetId, scope) {
        if (scope) throw new FormulaError('Scoped formula ranges are only supported inside table section rows.');
        return FormulaResolver.resolveRegisteredRange(
          registry,
          startId,
          endId,
          sheetId ?? currentSheetId,
          currentSheetId,
        );
      },
    };
  }

  static createTableContext(options: TableFormulaContextOptions): FormulaCompileContext {
    return {
      resolveCellId(id, sheetId) {
        if (sheetId && sheetId !== options.currentSheetId) {
          return FormulaResolver.resolveRegisteredCell(options.registry, id, sheetId, options.currentSheetId);
        }
        return FormulaResolver.resolveTableCellReference(id, options);
      },
      resolveRangeIds(startId, endId, sheetId, scope) {
        if (scope) return FormulaResolver.resolveScopedTableRange(startId, endId, sheetId, scope, options);
        if (sheetId && sheetId !== options.currentSheetId) {
          return FormulaResolver.resolveRegisteredRange(
            options.registry,
            startId,
            endId,
            sheetId,
            options.currentSheetId,
          );
        }
        return FormulaResolver.resolveTableRange(startId, endId, options);
      },
    };
  }

  // ─── Private helpers ──────────────────────────────────────────────────────

  private static resolveTableCellReference(id: string, options: TableFormulaContextOptions): string {
    const columnOffset = options.columnIdMap.get(id);

    if (columnOffset !== undefined) {
      return FormulaCompiler.formatCellAddress({ row: options.row, column: options.firstColumn + columnOffset });
    }

    const address = options.registry.resolve(id, options.currentSheetId);
    if (address) return FormulaCompiler.formatCellReference(address, options.currentSheetId);

    throw new FormulaError(`Formula references unknown cell id "${id}".`);
  }

  private static resolveTableRange(startId: string, endId: string, options: TableFormulaContextOptions): string {
    const startOffset = options.columnIdMap.get(startId);
    const endOffset = options.columnIdMap.get(endId);

    if (startOffset !== undefined && endOffset !== undefined) {
      FormulaResolver.assertRangeOrder(startOffset, endOffset);
      return `${FormulaCompiler.formatCellAddress({ row: options.row, column: options.firstColumn + startOffset })}:${FormulaCompiler.formatCellAddress({ row: options.row, column: options.firstColumn + endOffset })}`;
    }

    const startAddress = options.registry.resolve(startId, options.currentSheetId);
    const endAddress = options.registry.resolve(endId, options.currentSheetId);

    if (startAddress && endAddress) {
      FormulaResolver.assertAddressRangeOrder(startAddress, endAddress);
      return `${FormulaCompiler.formatCellReference(startAddress, options.currentSheetId)}:${FormulaCompiler.formatCellReference(endAddress, options.currentSheetId)}`;
    }

    if (startOffset === undefined && !startAddress) {
      throw new FormulaError(`Formula references unknown range start id "${startId}".`);
    }

    throw new FormulaError(`Formula references unknown range end id "${endId}".`);
  }

  private static resolveScopedTableRange(
    startId: string,
    endId: string,
    sheetId: string | undefined,
    scope: FormulaRangeScope,
    options: TableFormulaContextOptions,
  ): string {
    if (sheetId && sheetId !== options.currentSheetId) {
      throw new FormulaError('Scoped formula ranges must reference the current table sheet.');
    }

    const rows = options.scopedRows?.[scope];
    if (!rows) throw new FormulaError('Scoped formula ranges are only supported inside table section rows.');

    const startOffset = options.columnIdMap.get(startId);
    if (startOffset === undefined) throw new FormulaError(`Formula references unknown range start id "${startId}".`);

    const endOffset = options.columnIdMap.get(endId);
    if (endOffset === undefined) throw new FormulaError(`Formula references unknown range end id "${endId}".`);

    FormulaResolver.assertRangeOrder(startOffset, endOffset);

    if (rows.length === 0) return '0';

    return FormulaResolver.buildDiscontiguousRange(
      rows,
      options.firstColumn + startOffset,
      options.firstColumn + endOffset,
    );
  }

  private static resolveRegisteredCell(
    registry: AddressRegistry,
    id: string,
    targetSheetId: string,
    currentSheetId: string,
  ): string {
    const address = registry.resolve(id, targetSheetId);
    if (!address) throw new FormulaError(`Formula references unknown cell id "${id}".`);
    return FormulaCompiler.formatCellReference(address, currentSheetId);
  }

  private static resolveRegisteredRange(
    registry: AddressRegistry,
    startId: string,
    endId: string,
    targetSheetId: string,
    currentSheetId: string,
  ): string {
    const start = FormulaResolver.requireRangeAddress(registry.resolve(startId, targetSheetId), startId, 'start');
    const end = FormulaResolver.requireRangeAddress(registry.resolve(endId, targetSheetId), endId, 'end');
    FormulaResolver.assertAddressRangeOrder(start, end);
    return `${FormulaCompiler.formatCellReference(start, currentSheetId)}:${FormulaCompiler.formatCellReference(end, currentSheetId)}`;
  }

  private static requireRangeAddress(
    address: CellAddress | undefined,
    id: string,
    boundary: 'start' | 'end',
  ): CellAddress {
    if (!address) throw new FormulaError(`Formula references unknown range ${boundary} id "${id}".`);
    return address;
  }

  /**
   * Rows arrive in insertion order (ascending row number) because writeDataRows
   * iterates data sequentially. No sort needed.
   */
  private static buildDiscontiguousRange(rows: RenderedDataRow[], startCol: number, endCol: number): string {
    const segments: string[] = [];
    let segStart = rows[0]!.row;
    let segEnd = segStart;

    for (let i = 1; i < rows.length; i++) {
      const row = rows[i]!.row;
      if (row === segEnd + 1) {
        segEnd = row;
      } else {
        segments.push(FormulaResolver.formatRangeSegment(segStart, segEnd, startCol, endCol));
        segStart = row;
        segEnd = row;
      }
    }

    segments.push(FormulaResolver.formatRangeSegment(segStart, segEnd, startCol, endCol));
    return segments.join(',');
  }

  private static formatRangeSegment(startRow: number, endRow: number, startCol: number, endCol: number): string {
    return `${FormulaCompiler.formatCellAddress({ row: startRow, column: startCol })}:${FormulaCompiler.formatCellAddress({ row: endRow, column: endCol })}`;
  }

  private static assertRangeOrder(startOffset: number, endOffset: number): void {
    if (endOffset < startOffset) throw new FormulaError('Formula range end id must resolve after start id.');
  }

  private static assertAddressRangeOrder(start: CellAddress, end: CellAddress): void {
    if (end.row < start.row || end.column < start.column) {
      throw new FormulaError('Formula range end id must resolve after start id.');
    }
  }
}
