import type ExcelJS from 'exceljs';
import type { CellStyleDefinition } from '../types';
import { cloneStylePart } from '../styles/style-resolver';

export function writeCell(
  worksheet: ExcelJS.Worksheet,
  row: number,
  column: number,
  value: any, // value or formula object
  style: CellStyleDefinition | undefined,
): void {
  const cell = worksheet.getCell(row, column);

  // If value is a formula object, handle it specifically
  if (value !== undefined && value !== null && typeof value === 'object' && 'formula' in value) {
    cell.value =
      value.result !== undefined
        ? { formula: value.formula, result: value.result, date1904: false }
        : { formula: value.formula, date1904: false };
  } else {
    cell.value = value ?? null;
  }

  if (style) {
    cell.style = cloneStylePart(style) as Partial<ExcelJS.Style>;
  }
}

export function writeMerge(
  worksheet: ExcelJS.Worksheet,
  startRow: number,
  startColumn: number,
  endRow: number,
  endColumn: number,
): void {
  if (startRow === endRow && startColumn === endColumn) return;

  worksheet.mergeCells(startRow, startColumn, endRow, endColumn);

  const masterCell = worksheet.getCell(startRow, startColumn);
  const style = masterCell.style;

  if (style && style.border) {
    for (let r = startRow; r <= endRow; r++) {
      for (let c = startColumn; c <= endColumn; c++) {
        if (r === startRow && c === startColumn) continue;
        (worksheet.getCell(r, c) as ExcelJS.Cell).style = {
          ...style,
        } as Partial<ExcelJS.Style>;
      }
    }
  }
}
