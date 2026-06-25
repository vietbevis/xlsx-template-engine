import type { AddressRegistry } from './address-registry';
import { FormulaError } from './errors';
import { formatCellAddress, formatCellReference, type CellAddress, type FormulaCompileContext } from './formula-engine';

/**
 * FormulaCompileContext cho grid block — lookup từ registry toàn workbook
 * (hỗ trợ cross-sheet refs).
 */
export function createGridFormulaContext(registry: AddressRegistry, currentSheetId: string): FormulaCompileContext {
  return {
    resolveCellId(id, sheetId) {
      return resolveRegisteredCell(registry, id, sheetId ?? currentSheetId, currentSheetId);
    },
    resolveRangeIds(startId, endId, sheetId, scope) {
      if (scope) {
        throw new FormulaError('Scoped formula ranges are only supported inside table section rows.');
      }

      return resolveRegisteredRange(registry, startId, endId, sheetId ?? currentSheetId, currentSheetId);
    },
  };
}

/**
 * FormulaCompileContext cho một data row của table.
 * Resolve `ref` sang cùng hàng theo columnIdMap.
 * Cross-sheet refs vẫn lookup qua registry.
 */
export function createTableRowFormulaContext(
  columnIdMap: Map<string, number>,
  row: number,
  firstColumn: number,
  registry: AddressRegistry,
  currentSheetId: string,
): FormulaCompileContext {
  return {
    resolveCellId(id, sheetId) {
      if (sheetId && sheetId !== currentSheetId) {
        return resolveRegisteredCell(registry, id, sheetId, currentSheetId);
      }

      const columnOffset = columnIdMap.get(id);

      if (columnOffset === undefined) {
        const registeredAddress = registry.resolve(id, currentSheetId);

        if (registeredAddress) {
          return formatCellReference(registeredAddress, currentSheetId);
        }

        throw new FormulaError(`Formula references unknown cell id "${id}".`);
      }

      return formatCellAddress({ row, column: firstColumn + columnOffset });
    },

    resolveRangeIds(startId, endId, sheetId, scope) {
      if (scope) {
        throw new FormulaError('Scoped formula ranges are only supported inside table section rows.');
      }

      if (sheetId && sheetId !== currentSheetId) {
        return resolveRegisteredRange(registry, startId, endId, sheetId, currentSheetId);
      }

      const startOffset = columnIdMap.get(startId);
      const endOffset = columnIdMap.get(endId);

      if (startOffset === undefined || endOffset === undefined) {
        const registeredRange = resolveRegisteredLocalRangeIfComplete(
          registry,
          startId,
          endId,
          currentSheetId,
          currentSheetId,
        );

        if (registeredRange) {
          return registeredRange;
        }
      }

      if (startOffset === undefined) {
        throw new FormulaError(`Formula references unknown range start id "${startId}".`);
      }
      if (endOffset === undefined) throw new FormulaError(`Formula references unknown range end id "${endId}".`);

      if (endOffset < startOffset) {
        throw new FormulaError('Formula range end id must resolve after start id.');
      }

      return [
        formatCellAddress({ row, column: firstColumn + startOffset }),
        formatCellAddress({ row, column: firstColumn + endOffset }),
      ].join(':');
    },
  };
}

export interface RenderedDataRow {
  data: Record<string, unknown>;
  row: number;
}

/**
 * FormulaCompileContext cho table section row (header/footer của group và table).
 * Hỗ trợ `scope: 'currentRows' | 'allRows'` để build range toàn bộ data rows
 * thay vì chỉ cùng hàng.
 */
export function createTableSectionFormulaContext(
  columnIdMap: Map<string, number>,
  currentRows: RenderedDataRow[],
  allRows: RenderedDataRow[],
  row: number,
  firstColumn: number,
  registry: AddressRegistry,
  currentSheetId: string,
): FormulaCompileContext {
  const rowCtx = createTableRowFormulaContext(columnIdMap, row, firstColumn, registry, currentSheetId);

  return {
    resolveCellId: rowCtx.resolveCellId,

    resolveRangeIds(startId, endId, sheetId, scope) {
      if (!scope) {
        return rowCtx.resolveRangeIds(startId, endId, sheetId);
      }

      if (sheetId && sheetId !== currentSheetId) {
        throw new FormulaError('Scoped formula ranges must reference the current table sheet.');
      }

      const rows = scope === 'allRows' ? allRows : currentRows;
      const startOffset = columnIdMap.get(startId);
      const endOffset = columnIdMap.get(endId);

      if (startOffset === undefined) {
        throw new FormulaError(`Formula references unknown range start id "${startId}".`);
      }
      if (endOffset === undefined) throw new FormulaError(`Formula references unknown range end id "${endId}".`);
      if (endOffset < startOffset) throw new FormulaError('Formula range end id must resolve after start id.');

      if (rows.length === 0) return '0';

      return buildDiscontiguousRange(rows, firstColumn + startOffset, firstColumn + endOffset);
    },
  };
}

/**
 * Gộp danh sách row (có thể không liên tục) thành chuỗi range Excel.
 * Ví dụ: rows [2,3,5] → "A2:B3,A5:B5"
 */
function buildDiscontiguousRange(rows: RenderedDataRow[], startCol: number, endCol: number): string {
  const sorted = [...rows].sort((a, b) => a.row - b.row);
  const segments: string[] = [];
  let segStart = sorted[0]!.row;
  let segEnd = segStart;

  for (const { row } of sorted.slice(1)) {
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

function resolveRegisteredCell(
  registry: AddressRegistry,
  id: string,
  targetSheetId: string,
  currentSheetId: string,
): string {
  const address = registry.resolve(id, targetSheetId);

  if (!address) {
    throw new FormulaError(`Formula references unknown cell id "${id}".`);
  }

  return formatCellReference(address, currentSheetId);
}

function resolveRegisteredRange(
  registry: AddressRegistry,
  startId: string,
  endId: string,
  targetSheetId: string,
  currentSheetId: string,
): string {
  const start = requireRegisteredRangeAddress(registry.resolve(startId, targetSheetId), startId, 'start');
  const end = requireRegisteredRangeAddress(registry.resolve(endId, targetSheetId), endId, 'end');

  if (end.row < start.row || end.column < start.column) {
    throw new FormulaError('Formula range end id must resolve after start id.');
  }

  return `${formatCellReference(start, currentSheetId)}:${formatCellReference(end, currentSheetId)}`;
}

function resolveRegisteredLocalRangeIfComplete(
  registry: AddressRegistry,
  startId: string,
  endId: string,
  targetSheetId: string,
  currentSheetId: string,
): string | undefined {
  const start = registry.resolve(startId, targetSheetId);
  const end = registry.resolve(endId, targetSheetId);

  if (!start || !end) {
    return undefined;
  }

  if (end.row < start.row || end.column < start.column) {
    throw new FormulaError('Formula range end id must resolve after start id.');
  }

  return `${formatCellReference(start, currentSheetId)}:${formatCellReference(end, currentSheetId)}`;
}

function requireRegisteredRangeAddress(
  address: CellAddress | undefined,
  id: string,
  boundary: 'start' | 'end',
): CellAddress {
  if (!address) {
    throw new FormulaError(`Formula references unknown range ${boundary} id "${id}".`);
  }

  return address;
}
