import { ReportEngineError } from '../errors';

export function gridOccupancyKey(rowOffset: number, columnOffset: number): string {
  return `${rowOffset}:${columnOffset}`;
}

export function resolveColSpan(
  colSpan: number | 'remaining' | undefined,
  columnOffset: number,
  sheetColumnCount: number,
): number {
  return colSpan === 'remaining' ? Math.max(1, sheetColumnCount - columnOffset) : (colSpan ?? 1);
}

export function markGridOccupied(
  occupied: Set<string>,
  rowOffset: number,
  columnOffset: number,
  rowSpan: number,
  colSpan: number,
): void {
  for (let row = rowOffset; row < rowOffset + rowSpan; row += 1) {
    for (let col = columnOffset; col < columnOffset + colSpan; col += 1) {
      occupied.add(gridOccupancyKey(row, col));
    }
  }
}

export function assertGridCellDoesNotOverlap(
  occupied: Set<string>,
  rowOffset: number,
  columnOffset: number,
  rowSpan: number,
  colSpan: number,
): void {
  for (let row = rowOffset; row < rowOffset + rowSpan; row += 1) {
    for (let column = columnOffset; column < columnOffset + colSpan; column += 1) {
      if (occupied.has(gridOccupancyKey(row, column))) {
        throw new ReportEngineError('Grid cell merge ranges must not overlap.');
      }
    }
  }
}
