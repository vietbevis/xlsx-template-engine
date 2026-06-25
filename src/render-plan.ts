import type { CellStyleDefinition, CellValue, StyleRegistry, StyleValue, WorkbookMetadata } from './types';

export interface RenderPlan {
  metadata?: WorkbookMetadata;
  defaultStyle?: CellStyleDefinition;
  styles?: StyleRegistry;
  sheets: RenderPlanSheet[];
}

export interface RenderPlanSheet {
  id: string;
  name: string;
  views?: RenderSheetView[];
  rows: RenderRow[];
  merges: RenderMergeRange[];
  columnWidths: RenderColumnWidth[];
  columnVisibility: RenderColumnVisibility[];
  rowHeights: RenderRowHeight[];
  rowVisibility: RenderRowVisibility[];
}

export interface RenderSheetView {
  state: 'frozen';
  xSplit?: number;
  ySplit?: number;
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
  formulaResult?: CellValue;
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

export interface RenderColumnVisibility {
  column: number;
  hidden: boolean;
}

export interface RenderRowHeight {
  row: number;
  height: number;
}

export interface RenderRowVisibility {
  row: number;
  hidden: boolean;
}

export interface RenderLink {
  target: string;
  tooltip?: string;
}
