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
  namedRanges?: ResolvedNamedRange[];
  sheets: RenderPlanSheet[];
}

export interface ResolvedNamedRange {
  name: string;
  sheetName: string;
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
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

export type RenderCommand =
  | { type: 'cell'; sheetId: string; cell: RenderCell }
  | { type: 'merge'; sheetId: string; range: RenderMergeRange }
  | { type: 'columnWidth'; sheetId: string; width: RenderColumnWidth }
  | { type: 'columnVisibility'; sheetId: string; visibility: RenderColumnVisibility }
  | { type: 'rowVisibility'; sheetId: string; visibility: RenderRowVisibility }
  | { type: 'rowHeight'; sheetId: string; height: RenderRowHeight };
