import ExcelJS from 'exceljs';
import type { Writable } from 'stream';
import type { RenderCell, RenderPlan, RenderPlanSheet } from '../../compiler/render-plan';
import { ReportEngineError } from '../../core/errors';
import type { CellStyleDefinition, StyleValue } from '../../core/types';

interface StyleableWorksheet {
  getCell(row: number, column: number): ExcelJS.Cell;
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

        this.applyMergedStyleCoverage(sheet, sheetPlan, renderPlan, rowPlan.index);

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

      this.applyMergedStyleCoverage(sheet, sheetPlan, renderPlan);
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
    sheetPlan: RenderPlanSheet,
    renderPlan: RenderPlan,
    rowIndex?: number,
  ): void {
    for (const merge of sheetPlan.merges) {
      if (rowIndex !== undefined && (rowIndex < merge.startRow || rowIndex > merge.endRow)) {
        continue;
      }

      const masterCellPlan = this.findCellPlan(sheetPlan, merge.startRow, merge.startColumn);

      if (!masterCellPlan?.style && !masterCellPlan?.inlineStyle) {
        continue;
      }

      const style = this.resolveCellStyle(masterCellPlan, renderPlan);

      if (!style?.border) {
        continue;
      }

      const startRow = rowIndex ?? merge.startRow;
      const endRow = rowIndex ?? merge.endRow;

      for (let row = startRow; row <= endRow; row += 1) {
        for (let column = merge.startColumn; column <= merge.endColumn; column += 1) {
          sheet.getCell(row, column).style = cloneStylePart(style) as Partial<ExcelJS.Style>;
        }
      }
    }
  }

  private findCellPlan(
    sheetPlan: RenderPlanSheet,
    rowIndex: number,
    columnIndex: number,
  ): RenderCell | undefined {
    return sheetPlan.rows
      .find((row) => row.index === rowIndex)
      ?.cells.find((cell) => cell.column === columnIndex);
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
