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

interface StyledMergeCoverage {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  style: CellStyleDefinition;
}

/**
 * Writer ghi trực tiếp vào `ExcelJS.Worksheet`.
 * Scoped theo sheet: mỗi sheet có một SheetWriter riêng.
 *
 * Merge style coverage (propagating border styles to all cells in a merge range)
 * is applied on `finish()` after all cells and merges have been written.
 */
export class SheetWriter {
  private readonly merges: WriterMergeRange[] = [];
  /**
   * Track cells written to the master (top-left) position of each merge range
   * so we can resolve their styles for merge coverage in `finish()`.
   */
  private readonly masterCells = new Map<string, WriterCell>();

  constructor(
    private readonly worksheet: ExcelJS.Worksheet,
    private readonly config: SheetWriterConfig = {},
  ) {}

  addCell(cell: WriterCell): void {
    assertPositiveInteger(cell.row, 'cell row');
    assertPositiveInteger(cell.column, 'cell column');

    // Track cell for potential merge style coverage lookup
    this.masterCells.set(`${cell.row}:${cell.column}`, cell);

    const excelCell = this.worksheet.getRow(cell.row).getCell(cell.column);

    // Set value or formula
    excelCell.value = cell.formula
      ? this.createFormulaValue(cell.formula, cell.formulaResult ?? cell.value)
      : (cell.value ?? null);

    // Resolve and apply style
    const style = this.resolveCellStyle(cell);
    if (style) {
      // ExcelJS does not deep-clone on assignment, so we clone once here.
      excelCell.style = cloneStylePart(style) as Partial<ExcelJS.Style>;
    }
  }

  addMerge(range: WriterMergeRange): void {
    const normalized = normalizeMergeRange(range);
    if (normalized.type === 'skip-single-cell') return;
    this.merges.push(normalized.range);
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

  /**
   * Finalize the sheet: apply all merges and propagate border styles to
   * every cell within each merge range.
   */
  finish(): void {
    // Apply merges
    for (const merge of this.merges) {
      this.worksheet.mergeCells(merge.startRow, merge.startColumn, merge.endRow, merge.endColumn);
    }

    // Build and apply styled merge coverage
    const coverage = this.buildStyledMergeCoverage();
    for (const merge of coverage) {
      for (let row = merge.startRow; row <= merge.endRow; row++) {
        for (let column = merge.startColumn; column <= merge.endColumn; column++) {
          (this.worksheet.getCell(row, column) as ExcelJS.Cell).style = {
            ...merge.style,
          } as Partial<ExcelJS.Style>;
        }
      }
    }
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

  /**
   * Build list of merge ranges that carry border styles.
   * We look up the master cell (top-left) from the tracked masterCells map.
   */
  private buildStyledMergeCoverage(): StyledMergeCoverage[] {
    const coverage: StyledMergeCoverage[] = [];

    for (const merge of this.merges) {
      const masterCell = this.masterCells.get(`${merge.startRow}:${merge.startColumn}`);
      if (!masterCell) continue;

      const style = this.resolveCellStyle(masterCell);
      if (!style?.border) continue;

      coverage.push({ ...merge, style });
    }

    return coverage;
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

/**
 * Chuẩn hóa merge range:
 * - Validate tọa độ là số nguyên dương và end >= start.
 * - Nếu chỉ 1 ô → skip (không cần merge).
 * - Ngược lại → trả về range hợp lệ.
 */
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
