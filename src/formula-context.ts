import type { AddressRegistry } from './address-registry';
import { FormulaError } from './errors';
import { type CellAddress, formatCellAddress, formatCellReference, type FormulaCompileContext } from './formula-engine';
import type { FormulaRangeScope } from './types';

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

export function createGridFormulaContext(registry: AddressRegistry, currentSheetId: string): FormulaCompileContext {
  return {
    resolveCellId(id, sheetId) {
      return resolveRegisteredCell(registry, id, sheetId ?? currentSheetId, currentSheetId);
    },
    resolveRangeIds(startId, endId, sheetId, scope) {
      if (scope) throw new FormulaError('Scoped formula ranges are only supported inside table section rows.');
      return resolveRegisteredRange(registry, startId, endId, sheetId ?? currentSheetId, currentSheetId);
    },
  };
}

export function createTableFormulaContext(options: TableFormulaContextOptions): FormulaCompileContext {
  return {
    resolveCellId(id, sheetId) {
      if (sheetId && sheetId !== options.currentSheetId) {
        return resolveRegisteredCell(options.registry, id, sheetId, options.currentSheetId);
      }
      return resolveTableCellReference(id, options);
    },
    resolveRangeIds(startId, endId, sheetId, scope) {
      if (scope) return resolveScopedTableRange(startId, endId, sheetId, scope, options);
      if (sheetId && sheetId !== options.currentSheetId) {
        return resolveRegisteredRange(options.registry, startId, endId, sheetId, options.currentSheetId);
      }
      return resolveTableRange(startId, endId, options);
    },
  };
}

function resolveTableCellReference(id: string, options: TableFormulaContextOptions): string {
  const columnOffset = options.columnIdMap.get(id);

  if (columnOffset !== undefined) {
    return formatCellAddress({ row: options.row, column: options.firstColumn + columnOffset });
  }

  const address = options.registry.resolve(id, options.currentSheetId);
  if (address) return formatCellReference(address, options.currentSheetId);

  throw new FormulaError(`Formula references unknown cell id "${id}".`);
}

function resolveTableRange(startId: string, endId: string, options: TableFormulaContextOptions): string {
  const startOffset = options.columnIdMap.get(startId);
  const endOffset = options.columnIdMap.get(endId);

  if (startOffset !== undefined && endOffset !== undefined) {
    assertRangeOrder(startOffset, endOffset);
    return `${formatCellAddress({ row: options.row, column: options.firstColumn + startOffset })}:${formatCellAddress({ row: options.row, column: options.firstColumn + endOffset })}`;
  }

  const startAddress = options.registry.resolve(startId, options.currentSheetId);
  const endAddress = options.registry.resolve(endId, options.currentSheetId);

  if (startAddress && endAddress) {
    assertAddressRangeOrder(startAddress, endAddress);
    return `${formatCellReference(startAddress, options.currentSheetId)}:${formatCellReference(endAddress, options.currentSheetId)}`;
  }

  if (startOffset === undefined && !startAddress) {
    throw new FormulaError(`Formula references unknown range start id "${startId}".`);
  }

  throw new FormulaError(`Formula references unknown range end id "${endId}".`);
}

function resolveScopedTableRange(
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

  assertRangeOrder(startOffset, endOffset);

  if (rows.length === 0) return '0';

  return buildDiscontiguousRange(rows, options.firstColumn + startOffset, options.firstColumn + endOffset);
}

function resolveRegisteredCell(
  registry: AddressRegistry,
  id: string,
  targetSheetId: string,
  currentSheetId: string,
): string {
  const address = registry.resolve(id, targetSheetId);
  if (!address) throw new FormulaError(`Formula references unknown cell id "${id}".`);
  return formatCellReference(address, currentSheetId);
}

function resolveRegisteredRange(
  registry: AddressRegistry,
  startId: string,
  endId: string,
  targetSheetId: string,
  currentSheetId: string,
): string {
  const start = requireRangeAddress(registry.resolve(startId, targetSheetId), startId, 'start');
  const end = requireRangeAddress(registry.resolve(endId, targetSheetId), endId, 'end');
  assertAddressRangeOrder(start, end);
  return `${formatCellReference(start, currentSheetId)}:${formatCellReference(end, currentSheetId)}`;
}

function requireRangeAddress(address: CellAddress | undefined, id: string, boundary: 'start' | 'end'): CellAddress {
  if (!address) throw new FormulaError(`Formula references unknown range ${boundary} id "${id}".`);
  return address;
}

/**
 * Rows arrive in insertion order (ascending row number) because writeDataRows
 * iterates data sequentially. No sort needed.
 */
function buildDiscontiguousRange(rows: RenderedDataRow[], startCol: number, endCol: number): string {
  const segments: string[] = [];
  let segStart = rows[0]!.row;
  let segEnd = segStart;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i]!.row;
    if (row === segEnd + 1) {
      segEnd = row;
    } else {
      segments.push(formatRangeSegment(segStart, segEnd, startCol, endCol));
      segStart = row;
      segEnd = row;
    }
  }

  segments.push(formatRangeSegment(segStart, segEnd, startCol, endCol));
  return segments.join(',');
}

function formatRangeSegment(startRow: number, endRow: number, startCol: number, endCol: number): string {
  return `${formatCellAddress({ row: startRow, column: startCol })}:${formatCellAddress({ row: endRow, column: endCol })}`;
}

function assertRangeOrder(startOffset: number, endOffset: number): void {
  if (endOffset < startOffset) throw new FormulaError('Formula range end id must resolve after start id.');
}

function assertAddressRangeOrder(start: CellAddress, end: CellAddress): void {
  if (end.row < start.row || end.column < start.column) {
    throw new FormulaError('Formula range end id must resolve after start id.');
  }
}
