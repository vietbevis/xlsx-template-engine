import { ReportEngineError } from '../core/errors';

export class LayoutCursor {
  constructor(
    public row = 1,
    public column = 1,
  ) {
    assertPositiveInteger(row, 'row');
    assertPositiveInteger(column, 'column');
  }

  advanceRows(count = 1): void {
    assertPositiveInteger(count, 'row advance count');
    this.row += count;
  }

  advanceColumns(count = 1): void {
    assertPositiveInteger(count, 'column advance count');
    this.column += count;
  }

  clone(): LayoutCursor {
    return new LayoutCursor(this.row, this.column);
  }
}

function assertPositiveInteger(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new ReportEngineError(`Layout cursor ${label} must be a positive integer.`);
  }
}
