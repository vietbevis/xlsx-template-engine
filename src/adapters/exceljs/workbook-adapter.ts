import ExcelJS from 'exceljs';
import type { Writable } from 'stream';
import type { RenderCell, RenderPlan, RenderPlanSheet } from '../../compiler/render-plan';
import { ReportEngineError } from '../../core/errors';
import type { CellStyleDefinition, StyleValue } from '../../core/types';

interface StyleableWorksheet {
  getCell(row: number, column: number): ExcelJS.Cell;
}

type CellPlanIndex = Map<number, Map<number, RenderCell>>;

interface StyledMergeCoverage {
  startRow: number;
  endRow: number;
  startColumn: number;
  endColumn: number;
  style: CellStyleDefinition;
}

export class ExcelJsWorkbookAdapter {
  async writeFile(renderPlan: RenderPlan, filePath: string): Promise<void> {
    const workbook = this.createStreamingWorkbook(renderPlan, { filename: filePath });
    await workbook.commit();
  }

  async writeBuffer(renderPlan: RenderPlan): Promise<Buffer> {
    const workbook = this.createWorkbook(renderPlan);
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
  }

  async writeStream(renderPlan: RenderPlan, stream: Writable): Promise<void> {
    const workbook = this.createStreamingWorkbook(renderPlan, { stream });
    await workbook.commit();
  }

  private createStreamingWorkbook(
    renderPlan: RenderPlan,
    options: Partial<ExcelJS.stream.xlsx.WorkbookStreamWriterOptions>,
  ): ExcelJS.stream.xlsx.WorkbookWriter {
    const workbook = new ExcelJS.stream.xlsx.WorkbookWriter({
      ...options,
      useStyles: true,
    });

    if (renderPlan.metadata?.author) workbook.creator = renderPlan.metadata.author;

    for (const sheetPlan of renderPlan.sheets) {
      const sheet = workbook.addWorksheet(sheetPlan.name);
      const cellPlanIndex = this.createCellPlanIndex(sheetPlan);
      const mergeCoverage = this.createStyledMergeCoverage(sheetPlan, renderPlan, cellPlanIndex);
      const mergeCoverageByRow = this.indexMergeCoverageByRow(mergeCoverage);

      for (const columnWidth of sheetPlan.columnWidths) {
        sheet.getColumn(columnWidth.column).width = columnWidth.width;
      }

      for (const rowHeight of sheetPlan.rowHeights) {
        sheet.getRow(rowHeight.row).height = rowHeight.height;
      }

      for (const merge of sheetPlan.merges) {
        sheet.mergeCells(merge.startRow, merge.startColumn, merge.endRow, merge.endColumn);
      }

      for (const rowPlan of sheetPlan.rows) {
        const row = sheet.getRow(rowPlan.index);

        for (const cellPlan of rowPlan.cells) {
          this.applyCellPlan(row.getCell(cellPlan.column), cellPlan, renderPlan);
        }

        this.applyMergedStyleCoverage(
          sheet,
          mergeCoverageByRow.get(rowPlan.index) ?? [],
          rowPlan.index,
        );

        row.commit();
      }
    }

    return workbook;
  }

  private createWorkbook(renderPlan: RenderPlan): ExcelJS.Workbook {
    const workbook = new ExcelJS.Workbook();

    if (renderPlan.metadata?.author) workbook.creator = renderPlan.metadata.author;

    for (const sheetPlan of renderPlan.sheets) {
      const sheet = workbook.addWorksheet(sheetPlan.name);
      const cellPlanIndex = this.createCellPlanIndex(sheetPlan);
      const mergeCoverage = this.createStyledMergeCoverage(sheetPlan, renderPlan, cellPlanIndex);

      for (const columnWidth of sheetPlan.columnWidths) {
        sheet.getColumn(columnWidth.column).width = columnWidth.width;
      }

      for (const rowHeight of sheetPlan.rowHeights) {
        sheet.getRow(rowHeight.row).height = rowHeight.height;
      }

      for (const rowPlan of sheetPlan.rows) {
        const row = sheet.getRow(rowPlan.index);

        for (const cellPlan of rowPlan.cells) {
          this.applyCellPlan(row.getCell(cellPlan.column), cellPlan, renderPlan);
        }
      }

      for (const merge of sheetPlan.merges) {
        sheet.mergeCells(merge.startRow, merge.startColumn, merge.endRow, merge.endColumn);
      }

      this.applyMergedStyleCoverage(sheet, mergeCoverage);
    }

    return workbook;
  }

  private applyCellPlan(
    cell: ExcelJS.Cell,
    cellPlan: RenderPlan['sheets'][number]['rows'][number]['cells'][number],
    renderPlan: RenderPlan,
  ): void {
    cell.value = cellPlan.formula
      ? this.createFormulaValue(cellPlan.formula, cellPlan.value)
      : (cellPlan.value ?? null);

    const style = this.resolveCellStyle(cellPlan, renderPlan);

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
          sheet.getCell(row, column).style = cloneStylePart(merge.style) as Partial<ExcelJS.Style>;
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

  private createStyledMergeCoverage(
    sheetPlan: RenderPlanSheet,
    renderPlan: RenderPlan,
    cellPlanIndex: CellPlanIndex,
  ): StyledMergeCoverage[] {
    const coverage: StyledMergeCoverage[] = [];

    for (const merge of sheetPlan.merges) {
      const masterCellPlan = cellPlanIndex.get(merge.startRow)?.get(merge.startColumn);

      if (!masterCellPlan?.style && !masterCellPlan?.inlineStyle) {
        continue;
      }

      const style = this.resolveCellStyle(masterCellPlan, renderPlan);

      if (!style?.border) {
        continue;
      }

      coverage.push({
        ...merge,
        style,
      });
    }

    return coverage;
  }

  private indexMergeCoverageByRow(
    mergeCoverage: readonly StyledMergeCoverage[],
  ): Map<number, StyledMergeCoverage[]> {
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
    if (typeof result === 'string' || typeof result === 'number' || result instanceof Date) {
      return { formula, result, date1904: false };
    }

    return { formula, date1904: false };
  }

  private resolveCellStyle(
    cellPlan: RenderCell,
    renderPlan: RenderPlan,
  ): CellStyleDefinition | undefined {
    const baseStyle = this.resolveStyleValue(cellPlan.style, renderPlan);
    const style = this.mergeCellStyles(renderPlan.defaultStyle, baseStyle);

    return this.mergeCellStyles(style, cellPlan.inlineStyle);
  }

  private resolveStyleValue(
    style: StyleValue | undefined,
    renderPlan: RenderPlan,
  ): CellStyleDefinition | undefined {
    if (!style) {
      return undefined;
    }

    if (typeof style !== 'string') {
      return style;
    }

    const registryStyle = renderPlan.styles?.[style];

    if (!registryStyle) {
      throw new ReportEngineError(`Render plan references unknown style "${style}".`);
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

function cloneStylePart(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((item) => cloneStylePart(item));
  }

  if (!isPlainObject(value)) {
    return value;
  }

  return Object.fromEntries(
    Object.entries(value).map(([key, childValue]) => [key, cloneStylePart(childValue)]),
  );
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
