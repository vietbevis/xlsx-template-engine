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

export class RenderPlanBuilder {
  private readonly sheets = new Map<string, RenderPlanSheet>();
  private readonly mergeRanges: MergeRange[] = [];

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
    row.cells.sort((left, right) => left.column - right.column);
  }

  addMerge(sheetId: string, range: RenderMergeRange): void {
    const sheet = this.getSheet(sheetId);
    const normalizedRange = normalizeMergeRange(sheetId, range);

    if (!normalizedRange) {
      return;
    }

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
      metadata: this.metadata
        ? {
            ...this.metadata,
            keywords: this.metadata.keywords ? [...this.metadata.keywords] : undefined,
          }
        : undefined,
      styles: this.styles ? resolveStyleRegistry(this.styles) : undefined,
      sheets: Array.from(this.sheets.values()).map((sheet) => ({
        id: sheet.id,
        name: sheet.name,
        rows: sheet.rows.map((row) => ({
          index: row.index,
          cells: row.cells.map((cell) => ({
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

function resolveStyleRegistry(styles: StyleRegistry): StyleRegistry {
  const resolved = new Map<string, CellStyleDefinition>();

  for (const styleName of Object.keys(styles)) {
    resolved.set(styleName, resolveStyle(styleName, styles, resolved, new Set()));
  }

  return Object.fromEntries(resolved.entries());
}

function resolveStyle(
  styleName: string,
  styles: StyleRegistry,
  resolved: Map<string, CellStyleDefinition>,
  visiting: Set<string>,
): CellStyleDefinition {
  const existing = resolved.get(styleName);

  if (existing) {
    return existing;
  }

  const style = styles[styleName];

  if (!style) {
    throw new ReportEngineError(`Style "${styleName}" does not exist.`);
  }

  if (visiting.has(styleName)) {
    throw new ReportEngineError(`Style "${styleName}" has a circular extends chain.`);
  }

  visiting.add(styleName);

  const parentStyle = style.extends
    ? resolveStyle(style.extends, styles, resolved, visiting)
    : undefined;
  const mergedStyle = mergeStyles(parentStyle, style);

  visiting.delete(styleName);
  resolved.set(styleName, mergedStyle);

  return mergedStyle;
}

function mergeStyles(
  parent: CellStyleDefinition | undefined,
  style: CellStyleDefinition,
): CellStyleDefinition {
  const mergedStyle = mergeNested(parent, style) ?? {};
  delete mergedStyle.extends;
  return mergedStyle;
}

function mergeNested<T extends Record<string, unknown>>(
  parent: T | undefined,
  child: T | undefined,
): T | undefined {
  if (!parent && !child) {
    return undefined;
  }

  const merged: Record<string, unknown> = { ...(parent ?? {}) };

  for (const [key, value] of Object.entries(child ?? {})) {
    const parentValue = merged[key];

    if (isPlainObject(parentValue) && isPlainObject(value)) {
      merged[key] = mergeNested(parentValue, value);
      continue;
    }

    merged[key] = cloneStylePart(value);
  }

  return merged as T;
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
