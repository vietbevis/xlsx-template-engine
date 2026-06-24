import { ReportEngineError } from '../core/errors';
import type {
  CellStyleDefinition,
  StyleRegistry,
  StyleValue,
  WorkbookMetadata,
} from '../core/types';
import type {
  RenderCell,
  RenderColumnWidth,
  RenderMergeRange,
  RenderPlan,
  RenderPlanSheet,
  RenderRowHeight,
} from './render-plan';
import { assertMergeDoesNotOverlap, normalizeMergeRange, type MergeRange } from './merge-engine';

export interface RenderPlanBuilderOptions {
  metadata?: WorkbookMetadata;
  defaultStyle?: CellStyleDefinition;
  styles?: StyleRegistry;
}

export class RenderPlanBuilder {
  private readonly sheets = new Map<string, RenderPlanSheet>();
  private readonly rowsBySheet = new Map<string, Map<number, RenderPlanSheet['rows'][number]>>();
  private readonly mergeRanges: MergeRange[] = [];

  constructor(private readonly options: RenderPlanBuilderOptions = {}) {}

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
    this.rowsBySheet.set(id, new Map());
  }

  addCell(sheetId: string, cell: RenderCell): void {
    assertPositiveInteger(cell.row, 'cell row');
    assertPositiveInteger(cell.column, 'cell column');

    const sheet = this.getSheet(sheetId);
    const row = this.getOrCreateRow(sheet, cell.row);
    row.cells.push({
      ...cell,
      style: cloneStyleValue(cell.style),
      inlineStyle: cell.inlineStyle ? cloneStyle(cell.inlineStyle) : undefined,
      link: cell.link ? { ...cell.link } : undefined,
    });
  }

  addMerge(sheetId: string, range: RenderMergeRange): void {
    const sheet = this.getSheet(sheetId);
    const normalized = normalizeMergeRange(sheetId, range);

    if (normalized.type === 'skip-single-cell') {
      return;
    }

    const normalizedRange = normalized.range;
    assertMergeDoesNotOverlap(normalizedRange, this.mergeRanges);
    this.mergeRanges.push(normalizedRange);
    sheet.merges.push({
      startRow: normalizedRange.startRow,
      startColumn: normalizedRange.startColumn,
      endRow: normalizedRange.endRow,
      endColumn: normalizedRange.endColumn,
    });
  }

  setColumnWidth(sheetId: string, width: RenderColumnWidth): void {
    assertPositiveInteger(width.column, 'column width column');

    if (width.width <= 0) {
      throw new ReportEngineError('Column width must be greater than 0.');
    }

    this.getSheet(sheetId).columnWidths.push({ ...width });
  }

  setRowHeight(sheetId: string, height: RenderRowHeight): void {
    assertPositiveInteger(height.row, 'row height row');

    if (height.height <= 0) {
      throw new ReportEngineError('Row height must be greater than 0.');
    }

    this.getSheet(sheetId).rowHeights.push({ ...height });
  }

  build(): RenderPlan {
    return {
      metadata: this.options.metadata
        ? {
            ...this.options.metadata,
            keywords: this.options.metadata.keywords
              ? [...this.options.metadata.keywords]
              : undefined,
          }
        : undefined,
      defaultStyle: this.options.defaultStyle ? cloneStyle(this.options.defaultStyle) : undefined,
      styles: this.options.styles ? resolveStyleRegistry(this.options.styles) : undefined,
      sheets: Array.from(this.sheets.values()).map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        rows: [...sheet.rows].sort(compareRows).map((row) => ({
          index: row.index,
          cells: [...row.cells].sort(compareCells).map((cell) => ({
            ...cell,
            style: cloneStyleValue(cell.style),
            inlineStyle: cell.inlineStyle ? cloneStyle(cell.inlineStyle) : undefined,
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

  private getOrCreateRow(
    sheet: RenderPlanSheet,
    rowIndex: number,
  ): RenderPlanSheet['rows'][number] {
    const rowMap = this.rowsBySheet.get(sheet.id);

    if (!rowMap) {
      throw new ReportEngineError(`Render plan does not contain sheet "${sheet.id}".`);
    }

    const existingRow = rowMap.get(rowIndex);

    if (existingRow) {
      return existingRow;
    }

    const row = { index: rowIndex, cells: [] };
    sheet.rows.push(row);
    rowMap.set(rowIndex, row);
    return row;
  }
}

function compareRows(
  left: RenderPlanSheet['rows'][number],
  right: RenderPlanSheet['rows'][number],
): number {
  return left.index - right.index;
}

function compareCells(left: RenderCell, right: RenderCell): number {
  return left.column - right.column;
}

function assertPositiveInteger(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new ReportEngineError(`Render plan ${label} must be a positive integer.`);
  }
}

function resolveStyleRegistry(styles: StyleRegistry): StyleRegistry {
  return Object.fromEntries(
    Object.entries(styles).map(([styleName, style]) => [styleName, cloneStyle(style)]),
  );
}

function cloneStyleValue(style: StyleValue | undefined): StyleValue | undefined {
  if (typeof style === 'string' || style === undefined) {
    return style;
  }

  return cloneStyle(style);
}

function cloneStyle(style: CellStyleDefinition): CellStyleDefinition {
  return cloneStylePart(style) as CellStyleDefinition;
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
