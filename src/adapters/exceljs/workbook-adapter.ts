import ExcelJS from "exceljs";
import type { Writable } from "stream";
import type { RenderPlan } from "../../compiler/render-plan";
import { ReportEngineError } from "../../core/errors";
import type {
  BorderStyleDefinition,
  CellStyleDefinition,
  ColorStyleDefinition,
  FillStyleDefinition,
} from "../../core/types";

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
        sheet.mergeCells(
          merge.startRow,
          merge.startColumn,
          merge.endRow,
          merge.endColumn,
        );
      }

      for (const rowPlan of sheetPlan.rows) {
        const row = sheet.getRow(rowPlan.index);

        for (const cellPlan of rowPlan.cells) {
          this.applyCellPlan(row.getCell(cellPlan.column), cellPlan, renderPlan);
        }

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
        sheet.mergeCells(
          merge.startRow,
          merge.startColumn,
          merge.endRow,
          merge.endColumn,
        );
      }
    }

    return workbook;
  }

  private applyCellPlan(
    cell: ExcelJS.Cell,
    cellPlan: RenderPlan["sheets"][number]["rows"][number]["cells"][number],
    renderPlan: RenderPlan,
  ): void {
    cell.value = cellPlan.formula
      ? this.createFormulaValue(cellPlan.formula, cellPlan.value)
      : (cellPlan.value ?? null);

    if (cellPlan.style) {
      const style = renderPlan.styles?.[cellPlan.style];

      if (!style) {
        throw new ReportEngineError(`Render plan references unknown style "${cellPlan.style}".`);
      }

      cell.style = this.mapCellStyle(style);
    }
  }

  private createFormulaValue(formula: string, result: unknown): ExcelJS.CellValue {
    if (typeof result === "string" || typeof result === "number" || result instanceof Date) {
      return { formula, result, date1904: false };
    }

    return { formula, date1904: false };
  }

  private mapCellStyle(style: CellStyleDefinition): Partial<ExcelJS.Style> {
    return {
      font: style.font
        ? {
            name: style.font.name,
            size: style.font.size,
            bold: style.font.bold,
            italic: style.font.italic,
            underline: style.font.underline,
            color: style.font.color ? this.mapColor(style.font.color) : undefined,
          }
        : undefined,
      fill: style.fill ? this.mapFill(style.fill) : undefined,
      border: style.border ? this.mapBorder(style.border) : undefined,
      alignment: style.alignment
        ? {
            horizontal: style.alignment.horizontal,
            vertical: style.alignment.vertical,
            wrapText: style.alignment.wrapText,
          }
        : undefined,
      numFmt: style.numberFormat,
    };
  }

  private mapFill(fill: FillStyleDefinition): ExcelJS.Fill {
    return {
      type: "pattern",
      pattern: fill.pattern ?? "solid",
      fgColor: fill.foregroundColor ? this.mapColor(fill.foregroundColor) : { argb: "FFFFFFFF" },
      bgColor: fill.backgroundColor ? this.mapColor(fill.backgroundColor) : undefined,
    };
  }

  private mapBorder(border: BorderStyleDefinition): Partial<ExcelJS.Borders> {
    return {
      top: border.top ? this.mapBorderSide(border.top) : undefined,
      right: border.right ? this.mapBorderSide(border.right) : undefined,
      bottom: border.bottom ? this.mapBorderSide(border.bottom) : undefined,
      left: border.left ? this.mapBorderSide(border.left) : undefined,
    };
  }

  private mapBorderSide(side: NonNullable<BorderStyleDefinition[keyof BorderStyleDefinition]>): Partial<ExcelJS.Border> {
    return {
      style: side.style,
      color: side.color ? this.mapColor(side.color) : undefined,
    };
  }

  private mapColor(color: ColorStyleDefinition): Partial<ExcelJS.Color> {
    return { argb: color.argb.toUpperCase() };
  }
}
