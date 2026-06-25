import type { Block, GridRow, WorkbookDefinition } from '../types';
import { assertNever } from './common';
import { flattenColumns } from './table';

export function createSheetColumnCounts(workbook: WorkbookDefinition): Map<string, number> {
  const widths = new Map<string, number>();

  for (const sheet of workbook.sheets) {
    widths.set(sheet.id, Math.max(1, ...sheet.blocks.map(measureBlockColumnCount)));
  }

  return widths;
}

export function measureBlockColumnCount(block: Block): number {
  switch (block.type) {
    case 'title':
    case 'text':
      return block.colSpan === 'remaining' ? 1 : (block.colSpan ?? 1);
    case 'spacer':
    case 'divider':
      return 1;
    case 'grid':
      return Math.max(1, ...block.rows.map(measureGridRowColumnCount));
    case 'table':
    case 'table-groups':
      return flattenColumns(block.columns).length;
    default:
      return assertNever(block);
  }
}

export function measureGridRowColumnCount(row: GridRow): number {
  return row.cells.reduce((width, cell) => width + (cell.colSpan === 'remaining' ? 1 : (cell.colSpan ?? 1)), 0);
}
