import type {
  CellStyleDefinition,
  CellValue,
  StyleValue,
  StyleRegistry,
  WorkbookMetadata,
} from '../core/types';

export interface RenderPlan {
  metadata?: WorkbookMetadata;
  defaultStyle?: CellStyleDefinition;
  styles?: StyleRegistry;
  sheets: RenderPlanSheet[];
}

export interface RenderPlanSheet {
  id: string;
  name: string;
  rows: RenderRow[];
  merges: RenderMergeRange[];
  columnWidths: RenderColumnWidth[];
  rowHeights: RenderRowHeight[];
}

export interface RenderRow {
  index: number;
  cells: RenderCell[];
}

export interface RenderCell {
  row: number;
  column: number;
  value?: CellValue;
  formula?: string;
  link?: RenderLink;
  style?: StyleValue;
  inlineStyle?: CellStyleDefinition;
}

export interface RenderMergeRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

export interface RenderColumnWidth {
  column: number;
  width: number;
}

export interface RenderRowHeight {
  row: number;
  height: number;
}

export interface RenderLink {
  target: string;
  tooltip?: string;
}

export type RenderCommand =
  | { type: 'cell'; sheetId: string; cell: RenderCell }
  | { type: 'merge'; sheetId: string; range: RenderMergeRange }
  | { type: 'columnWidth'; sheetId: string; width: RenderColumnWidth }
  | { type: 'rowHeight'; sheetId: string; height: RenderRowHeight };
