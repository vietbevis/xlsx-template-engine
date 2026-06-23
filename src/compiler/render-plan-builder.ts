import { ReportEngineError } from "../core/errors";
import type { StyleRegistry, WorkbookMetadata } from "../core/types";
import type {
  RenderCell,
  RenderColumnWidth,
  RenderMergeRange,
  RenderPlan,
  RenderPlanSheet,
  RenderRowHeight,
} from "./render-plan";

export class RenderPlanBuilder {
  private readonly sheets = new Map<string, RenderPlanSheet>();

  constructor(
    private readonly metadata?: WorkbookMetadata,
    private readonly styles?: StyleRegistry,
  ) {}

  addSheet(id: string, name: string): void {
    if (this.sheets.has(id)) {
      throw new ReportEngineError(`Render plan already contains sheet "${id}".`);
    }

    this.sheets.set(id, {
      id,
      name,
      rows: [],
      merges: [],
      columnWidths: [],
      rowHeights: [],
    });
  }

  addCell(sheetId: string, cell: RenderCell): void {
    assertPositiveInteger(cell.row, "cell row");
    assertPositiveInteger(cell.column, "cell column");

    const sheet = this.getSheet(sheetId);
    const row = this.getOrCreateRow(sheet, cell.row);
    row.cells.push({ ...cell, link: cell.link ? { ...cell.link } : undefined });
    row.cells.sort((left, right) => left.column - right.column);
  }

  addMerge(sheetId: string, range: RenderMergeRange): void {
    assertPositiveInteger(range.startRow, "merge start row");
    assertPositiveInteger(range.startColumn, "merge start column");
    assertPositiveInteger(range.endRow, "merge end row");
    assertPositiveInteger(range.endColumn, "merge end column");

    if (range.endRow < range.startRow || range.endColumn < range.startColumn) {
      throw new ReportEngineError("Merge range end must be greater than or equal to start.");
    }

    this.getSheet(sheetId).merges.push({ ...range });
  }

  setColumnWidth(sheetId: string, width: RenderColumnWidth): void {
    assertPositiveInteger(width.column, "column width column");

    if (width.width <= 0) {
      throw new ReportEngineError("Column width must be greater than 0.");
    }

    this.getSheet(sheetId).columnWidths.push({ ...width });
  }

  setRowHeight(sheetId: string, height: RenderRowHeight): void {
    assertPositiveInteger(height.row, "row height row");

    if (height.height <= 0) {
      throw new ReportEngineError("Row height must be greater than 0.");
    }

    this.getSheet(sheetId).rowHeights.push({ ...height });
  }

  build(): RenderPlan {
    return {
      metadata: this.metadata
        ? {
            ...this.metadata,
            keywords: this.metadata.keywords ? [...this.metadata.keywords] : undefined,
          }
        : undefined,
      styles: this.styles ? { ...this.styles } : undefined,
      sheets: Array.from(this.sheets.values()).map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        rows: sheet.rows.map((row) => ({
          index: row.index,
          cells: row.cells.map((cell) => ({
            ...cell,
            link: cell.link ? { ...cell.link } : undefined,
          })),
        })),
        merges: sheet.merges.map((merge) => ({ ...merge })),
        columnWidths: sheet.columnWidths.map((width) => ({ ...width })),
        rowHeights: sheet.rowHeights.map((height) => ({ ...height })),
      })),
    };
  }

  private getSheet(sheetId: string): RenderPlanSheet {
    const sheet = this.sheets.get(sheetId);

    if (!sheet) {
      throw new ReportEngineError(`Render plan does not contain sheet "${sheetId}".`);
    }

    return sheet;
  }

  private getOrCreateRow(sheet: RenderPlanSheet, rowIndex: number) {
    const existingRow = sheet.rows.find((row) => row.index === rowIndex);

    if (existingRow) {
      return existingRow;
    }

    const row = { index: rowIndex, cells: [] };
    sheet.rows.push(row);
    sheet.rows.sort((left, right) => left.index - right.index);
    return row;
  }
}

function assertPositiveInteger(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new ReportEngineError(`Render plan ${label} must be a positive integer.`);
  }
}
