import ExcelJS from 'exceljs';
import { RenderError, ReportEngineError } from './errors';
import { assertPositiveInteger } from './helpers/common';
import { cloneStylePart, isPlainObject } from './helpers/utils';
import type {
  CellStyleDefinition,
  StyleRegistry,
  StyleValue,
  WriterCell,
  WriterColumnVisibility,
  WriterColumnWidth,
  WriterMergeRange,
  WriterRowHeight,
  WriterRowVisibility,
} from './types';

type FormulaResult = string | number | boolean | Date | ExcelJS.CellErrorValue;

export interface SheetWriterConfig {
  defaultStyle?: CellStyleDefinition;
  styles?: StyleRegistry;
}

/**
 * Writer ghi trực tiếp vào `ExcelJS.Worksheet`.
 * Scoped theo sheet: mỗi sheet có một SheetWriter riêng.
 *
 * Buffer-free: các thay đổi được ghi vào worksheet ngay lập tức.
 */
export class SheetWriter {
  constructor(
    private readonly worksheet: ExcelJS.Worksheet,
    private readonly config: SheetWriterConfig = {},
  ) {}

  addCell(cell: WriterCell): void {
    assertPositiveInteger(cell.row, 'cell row');
    assertPositiveInteger(cell.column, 'cell column');

    const excelCell = this.worksheet.getRow(cell.row).getCell(cell.column);

    excelCell.value = cell.formula
      ? this.createFormulaValue(cell.formula, cell.formulaResult ?? cell.value)
      : (cell.value ?? null);

    const style = this.resolveCellStyle(cell);
    if (style) {
      excelCell.style = cloneStylePart(style) as Partial<ExcelJS.Style>;
    }
  }

  addMerge(range: WriterMergeRange): void {
    const normalized = normalizeMergeRange(range);
    if (normalized.type === 'skip-single-cell') return;

    const { startRow, startColumn, endRow, endColumn } = normalized.range;

    this.worksheet.mergeCells(startRow, startColumn, endRow, endColumn);

    const masterCell = this.worksheet.getCell(startRow, startColumn);
    const style = masterCell.style;

    // Apply the master cell's border style to the rest of the merged cells
    if (style && style.border) {
      for (let r = startRow; r <= endRow; r++) {
        for (let c = startColumn; c <= endColumn; c++) {
          if (r === startRow && c === startColumn) continue;
          (this.worksheet.getCell(r, c) as ExcelJS.Cell).style = {
            ...style,
          } as Partial<ExcelJS.Style>;
        }
      }
    }
  }

  setColumnWidth(width: WriterColumnWidth): void {
    assertPositiveInteger(width.column, 'column width column');
    if (width.width < 0) throw new ReportEngineError('Column width must be greater than 0.');
    (this.worksheet.getColumn(width.column) as ExcelJS.Column).width = width.width;
  }

  setColumnHidden(visibility: WriterColumnVisibility): void {
    assertPositiveInteger(visibility.column, 'column visibility column');
    (this.worksheet.getColumn(visibility.column) as ExcelJS.Column).hidden = visibility.hidden;
  }

  setRowHeight(height: WriterRowHeight): void {
    assertPositiveInteger(height.row, 'row height row');
    if (height.height < 0) throw new ReportEngineError('Row height must be greater than 0.');
    this.worksheet.getRow(height.row).height = height.height;
  }

  setRowHidden(visibility: WriterRowVisibility): void {
    assertPositiveInteger(visibility.row, 'row visibility row');
    this.worksheet.getRow(visibility.row).hidden = visibility.hidden;
  }

  // ─── Private helpers ──────────────────────────────────────────────────────────

  private createFormulaValue(formula: string, result: unknown): ExcelJS.CellValue {
    return result !== undefined && result !== null
      ? { formula, result: result as FormulaResult, date1904: false }
      : { formula, date1904: false };
  }

  private resolveCellStyle(cell: WriterCell): CellStyleDefinition | undefined {
    const baseStyle = this.resolveStyleValue(cell.style);
    return mergeCellStyles(mergeCellStyles(this.config.defaultStyle, baseStyle), cell.inlineStyle);
  }

  private resolveStyleValue(style: StyleValue | undefined): CellStyleDefinition | undefined {
    if (!style) return undefined;
    if (typeof style !== 'string') return style;

    const registryStyle = this.config.styles?.[style];
    if (!registryStyle) throw new RenderError(`Unknown style "${style}".`);
    return registryStyle;
  }
}

// ─── Pure helpers ─────────────────────────────────────────────────────────────

function mergeCellStyles(
  base: CellStyleDefinition | undefined,
  override: CellStyleDefinition | undefined,
): CellStyleDefinition | undefined {
  if (!base && !override) return undefined;
  return mergeStylePart(base, override) ?? {};
}

function mergeStylePart<T extends Record<string, unknown>>(
  base: T | undefined,
  override: T | undefined,
): T | undefined {
  if (!base && !override) return undefined;

  const merged: Record<string, unknown> = { ...(base ?? {}) };

  for (const [key, value] of Object.entries(override ?? {})) {
    const baseValue = merged[key];
    merged[key] =
      isPlainObject(baseValue) && isPlainObject(value) ? mergeStylePart(baseValue, value) : cloneStylePart(value);
  }

  return merged as T;
}

type NormalizedMergeRange = { type: 'skip-single-cell' } | { type: 'range'; range: WriterMergeRange };

function normalizeMergeRange(range: WriterMergeRange): NormalizedMergeRange {
  assertPositiveInteger(range.startRow, 'Render plan merge start row');
  assertPositiveInteger(range.startColumn, 'Render plan merge start column');
  assertPositiveInteger(range.endRow, 'Render plan merge end row');
  assertPositiveInteger(range.endColumn, 'Render plan merge end column');

  if (range.endRow < range.startRow || range.endColumn < range.startColumn) {
    throw new ReportEngineError('Merge range end must be greater than or equal to start.');
  }

  if (range.startRow === range.endRow && range.startColumn === range.endColumn) {
    return { type: 'skip-single-cell' };
  }

  return { type: 'range', range };
}
