import ExcelJS from "exceljs";
import type { Writable } from "stream";
import type { RenderPlan } from "../../compiler/render-plan";

export class ExcelJsWorkbookAdapter {
  async writeFile(renderPlan: RenderPlan, filePath: string): Promise<void> {
    const workbook = this.createWorkbook(renderPlan);
    await workbook.xlsx.writeFile(filePath);
  }

  async writeBuffer(renderPlan: RenderPlan): Promise<Buffer> {
    const workbook = this.createWorkbook(renderPlan);
    const buffer = await workbook.xlsx.writeBuffer();
    return Buffer.from(buffer);
  }

  async writeStream(renderPlan: RenderPlan, stream: Writable): Promise<void> {
    const workbook = this.createWorkbook(renderPlan);
    await workbook.xlsx.write(stream);
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
          const cell = row.getCell(cellPlan.column);
          cell.value = cellPlan.formula
            ? this.createFormulaValue(cellPlan.formula, cellPlan.value)
            : (cellPlan.value ?? null);
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

  private createFormulaValue(formula: string, result: unknown): ExcelJS.CellValue {
    if (typeof result === "string" || typeof result === "number" || result instanceof Date) {
      return { formula, result, date1904: false };
    }

    return { formula, date1904: false };
  }
}
