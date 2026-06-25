import ExcelJS from 'exceljs';
import type { Writable } from 'stream';
import type { RenderCell, RenderPlan, RenderPlanSheet } from './render-plan';
import { RenderError, ReportEngineError } from './errors';
import type { CellStyleDefinition, StyleValue } from './types';
import { cloneStylePart, isPlainObject } from './helpers/utils';

interface StyleableWorksheet {
  getCell(row: number, column: number): ExcelJS.Cell;
}

interface RenderableWorksheet extends StyleableWorksheet {
  getColumn(column: number): ExcelJS.Column;
  getRow(row: number): ExcelJS.Row;
  mergeCells(startRow: number, startColumn: number, endRow: number, endColumn: number): void;
}

type CellPlanIndex = Map<number, Map<number, RenderCell>>;
type FormulaResult = string | number | boolean | Date | ExcelJS.CellErrorValue;

interface StyledMergeCoverage {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  style: CellStyleDefinition;
}

export class ExcelJsWorkbookRenderer {
  constructor(private readonly renderPlan: RenderPlan) {}

  async writeFile(filePath: string): Promise<void> {
    try {
      const workbook = this.createStreamingWorkbook({ filename: filePath });
      await workbook.commit();
    } catch (error) {
      throw normalizeRenderError(error);
    }
  }

  async writeBuffer(): Promise<Buffer> {
    try {
      const workbook = this.createWorkbook();
      const buffer = await workbook.xlsx.writeBuffer();
      return Buffer.from(buffer);
    } catch (error) {
      throw normalizeRenderError(error);
    }
  }

  async writeStream(stream: Writable): Promise<void> {
    try {
      const workbook = this.createStreamingWorkbook({ stream });
      await workbook.commit();
    } catch (error) {
      throw normalizeRenderError(error);
    }
  }

  private createStreamingWorkbook(
    options: Partial<ExcelJS.stream.xlsx.WorkbookStreamWriterOptions>,
  ): ExcelJS.stream.xlsx.WorkbookWriter {
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      ...options,
      useStyles: true,
    });

    if (this.renderPlan.metadata?.author) workbook.creator = this.renderPlan.metadata.author;

    for (const sheetPlan of this.renderPlan.sheets) {
      const sheet = workbook.addWorksheet(sheetPlan.name, {
        views: sheetPlan.views,
      });

      this.renderSheet(sheet, sheetPlan, { streaming: true });
    }

    return workbook;
  }

  private createWorkbook(): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();

    if (this.renderPlan.metadata?.author) workbook.creator = this.renderPlan.metadata.author;

    for (const sheetPlan of this.renderPlan.sheets) {
      const sheet = workbook.addWorksheet(sheetPlan.name, {
        views: sheetPlan.views,
      });

      this.renderSheet(sheet, sheetPlan, { streaming: false });
    }

    return workbook;
  }

  private renderSheet(sheet: RenderableWorksheet, sheetPlan: RenderPlanSheet, options: { streaming: boolean }): void {
    const mergeCoverage = this.createStyledMergeCoverage(sheetPlan, this.createCellPlanIndex(sheetPlan));
    const mergeCoverageByRow = options.streaming ? this.indexMergeCoverageByRow(mergeCoverage) : undefined;

    this.applySheetLayout(sheet, sheetPlan);

    if (options.streaming) {
      this.applySheetMerges(sheet, sheetPlan);
    }

    for (const rowPlan of sheetPlan.rows) {
      const row = sheet.getRow(rowPlan.index);

      for (const cellPlan of rowPlan.cells) {
        this.applyCellPlan(row.getCell(cellPlan.column), cellPlan);
      }

      if (options.streaming) {
        // Streaming rows cannot be edited after commit.
        this.applyMergedStyleCoverage(sheet, mergeCoverageByRow?.get(rowPlan.index) ?? [], rowPlan.index);
        row.commit();
      }
    }

    if (!options.streaming) {
      this.applySheetMerges(sheet, sheetPlan);
      this.applyMergedStyleCoverage(sheet, mergeCoverage);
    }
  }

  private applySheetLayout(sheet: RenderableWorksheet, sheetPlan: RenderPlanSheet): void {
    for (const columnWidth of sheetPlan.columnWidths) {
      sheet.getColumn(columnWidth.column).width = columnWidth.width;
    }

    for (const columnVisibility of sheetPlan.columnVisibility ?? []) {
      sheet.getColumn(columnVisibility.column).hidden = columnVisibility.hidden;
    }

    for (const rowHeight of sheetPlan.rowHeights) {
      sheet.getRow(rowHeight.row).height = rowHeight.height;
    }

    for (const rowVisibility of sheetPlan.rowVisibility ?? []) {
      sheet.getRow(rowVisibility.row).hidden = rowVisibility.hidden;
    }
  }

  private applySheetMerges(sheet: RenderableWorksheet, sheetPlan: RenderPlanSheet): void {
    for (const merge of sheetPlan.merges) {
      sheet.mergeCells(merge.startRow, merge.startColumn, merge.endRow, merge.endColumn);
    }
  }

  private applyCellPlan(cell: ExcelJS.Cell, cellPlan: RenderCell): void {
    cell.value = cellPlan.formula
      ? this.createFormulaValue(cellPlan.formula, cellPlan.formulaResult ?? cellPlan.value)
      : (cellPlan.value ?? null);

    const style = this.resolveCellStyle(cellPlan);

    if (style) {
      cell.style = cloneStylePart(style) as Partial<ExcelJS.Style>;
    }
  }

  private applyMergedStyleCoverage(
    sheet: StyleableWorksheet,
    mergeCoverage: readonly StyledMergeCoverage[],
    rowIndex?: number,
  ): void {
    for (const merge of mergeCoverage) {
      const startRow = rowIndex ?? merge.startRow;
      const endRow = rowIndex ?? merge.endRow;

      for (let row = startRow; row <= endRow; row += 1) {
        for (let column = merge.startColumn; column <= merge.endColumn; column += 1) {
          sheet.getCell(row, column).style = { ...merge.style } as Partial<ExcelJS.Style>;
        }
      }
    }
  }

  private createCellPlanIndex(sheetPlan: RenderPlanSheet): CellPlanIndex {
    const index: CellPlanIndex = new Map();

    for (const row of sheetPlan.rows) {
      const cellsByColumn = new Map<number, RenderCell>();

      for (const cell of row.cells) {
        cellsByColumn.set(cell.column, cell);
      }

      index.set(row.index, cellsByColumn);
    }

    return index;
  }

  private createStyledMergeCoverage(sheetPlan: RenderPlanSheet, cellPlanIndex: CellPlanIndex): StyledMergeCoverage[] {
    const coverage: StyledMergeCoverage[] = [];

    for (const merge of sheetPlan.merges) {
      const masterCellPlan = cellPlanIndex.get(merge.startRow)?.get(merge.startColumn);

      if (!masterCellPlan?.style && !masterCellPlan?.inlineStyle) {
        continue;
      }

      const style = this.resolveCellStyle(masterCellPlan);

      if (!style?.border) {
        continue;
      }

      coverage.push({ ...merge, style });
    }

    return coverage;
  }

  private indexMergeCoverageByRow(mergeCoverage: readonly StyledMergeCoverage[]): Map<number, StyledMergeCoverage[]> {
    const index = new Map<number, StyledMergeCoverage[]>();

    for (const merge of mergeCoverage) {
      for (let row = merge.startRow; row <= merge.endRow; row += 1) {
        const rowMerges = index.get(row) ?? [];
        rowMerges.push(merge);
        index.set(row, rowMerges);
      }
    }

    return index;
  }

  private createFormulaValue(formula: string, result: unknown): ExcelJS.CellValue {
    if (result !== undefined && result !== null) {
      return {
        formula,
        result: result as FormulaResult,
        date1904: false,
      };
    }

    return { formula, date1904: false };
  }

  private resolveCellStyle(cellPlan: RenderCell): CellStyleDefinition | undefined {
    const baseStyle = this.resolveStyleValue(cellPlan.style);
    const style = this.mergeCellStyles(this.renderPlan.defaultStyle, baseStyle);

    return this.mergeCellStyles(style, cellPlan.inlineStyle);
  }

  private resolveStyleValue(style: StyleValue | undefined): CellStyleDefinition | undefined {
    if (!style) {
      return undefined;
    }

    if (typeof style !== 'string') {
      return style;
    }

    const registryStyle = this.renderPlan.styles?.[style];

    if (!registryStyle) {
      throw new RenderError(`Render plan references unknown style "${style}".`);
    }

    return registryStyle;
  }

  private mergeCellStyles(
    base: CellStyleDefinition | undefined,
    override: CellStyleDefinition | undefined,
  ): CellStyleDefinition | undefined {
    if (!base && !override) {
      return undefined;
    }

    return this.mergeStylePart(base, override) ?? {};
  }

  private mergeStylePart<T extends Record<string, unknown>>(
    base: T | undefined,
    override: T | undefined,
  ): T | undefined {
    if (!base && !override) {
      return undefined;
    }

    const merged: Record<string, unknown> = { ...(base ?? {}) };

    for (const [key, value] of Object.entries(override ?? {})) {
      const baseValue = merged[key];

      if (isPlainObject(baseValue) && isPlainObject(value)) {
        merged[key] = this.mergeStylePart(baseValue, value);
        continue;
      }

      merged[key] = cloneStylePart(value);
    }

    return merged as T;
  }
}

function normalizeRenderError(error: unknown): Error {
  if (error instanceof RenderError) {
    return error;
  }

  if (error instanceof ReportEngineError) {
    return new RenderError(error.message);
  }

  return error instanceof Error ? error : new RenderError(String(error));
}
