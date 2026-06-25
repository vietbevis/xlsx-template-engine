import type ExcelJS from 'exceljs';

export type CellValue = Exclude<ExcelJS.CellValue, undefined>;

export type CellContent = CellValue | FormulaDefinition;

/** Column node lá (không có children) trong cây cột của table. */
export type TableColumnNode = Extract<Block, { type: 'table' }>['columns'][number];
export type TableLeafColumn = TableColumnNode & { children?: undefined };

export type FormulaDefinition = FormulaTemplateDefinition | RefFormulaDefinition | RangeFormulaDefinition;

export interface FormulaTemplateDefinition {
  type: 'formula_template';
  strings: readonly string[];
  exprs: readonly FormulaExpression[];
}

export type FormulaExpression = string | number | boolean | null | FormulaDefinition;

export interface RefFormulaDefinition {
  type: 'ref';
  id: string;
  sheetId?: string;
}

export interface RangeFormulaDefinition {
  type: 'range';
  startId: string;
  endId: string;
  sheetId?: string;
  scope?: FormulaRangeScope;
}

export type FormulaRangeScope = 'currentRows' | 'allRows';

export type StyleRegistry = Record<string, CellStyleDefinition>;

export type StyleValue<TStyleName extends string = string> = TStyleName | CellStyleDefinition;

export interface StyleReference<TStyleName extends string = string> {
  style?: StyleValue<TStyleName>;
}

export type CellStyleDefinition = Partial<ExcelJS.Style>;

export interface WorkbookMetadata {
  title?: string;
  author?: string;
  company?: string;
  subject?: string;
  keywords?: string[];
}

export interface WorkbookDefinition {
  metadata?: WorkbookMetadata;
  defaultStyle?: CellStyleDefinition;
  styles?: StyleRegistry;
  sheets: readonly SheetDefinition[];
}

export interface SheetDefinition {
  id: string;
  name: string;
  freezePane?: SheetFreezePane;
  blocks: readonly Block[];
}

export interface SheetFreezePane {
  rows?: number;
  columns?: number;
}

export type Block = GridBlock | TableBlock;

export interface BaseBlock {
  type: string;
}

export interface GridBlock extends BaseBlock {
  type: 'grid';
  rows: readonly GridRow[];
}

export interface GridRow {
  height?: number;
  cells: readonly GridCell[];
}

export interface GridCell extends StyleReference {
  id?: string;
  value?: CellContent;
  formulaResult?: CellValue;
  styleResolver?: (value: CellValue | undefined) => StyleValue | undefined;
  colSpan?: number | 'remaining';
  rowSpan?: number;
  width?: number;
}

export interface BaseTableBlock<Row = Record<string, unknown>> extends BaseBlock {
  columns: readonly TableColumn<Row>[];
  headerRowHeights?: readonly number[];
  bodyRowHeight?: number;
  headerStyle?: StyleValue;
  bodyStyle?: StyleValue;
  evenRowStyle?: StyleValue;
  oddRowStyle?: StyleValue;
  footerRows?: readonly TableFooterRow<Row>[];
  summaryStyle?: StyleValue;
  rowHidden?: (row: Row, index: number) => boolean;
  border?: TableBorderDefinition;
}

export interface TableBlock<Row = Record<string, unknown>> extends BaseTableBlock<Row> {
  type: 'table';
  data?: readonly Row[];
  groups?: readonly TableGroup<Row>[];
}

export interface TableGroup<Row = Record<string, unknown>> {
  headerRows?: readonly TableSectionRow<Row>[];
  data: readonly Row[];
  footerRows?: readonly TableSectionRow<Row>[];
}

export type TableBorderDefinition = ExcelJS.BorderStyle | Partial<ExcelJS.Borders>;

export interface TableSectionRow<Row = Record<string, unknown>> extends StyleReference {
  resetRows?: boolean;
  hidden?: boolean;
  height?: number;
  cells: readonly TableSectionCell<Row>[];
}

export interface TableSectionCell<Row = Record<string, unknown>> extends StyleReference {
  id?: string;
  column?: number;
  columnId?: string;
  value?: CellContent | TableSectionCellAccessor<Row>;
  formulaResult?: CellValue;
  styleResolver?: (value: CellValue | undefined) => StyleValue | undefined;
  colSpan?: number | 'remaining';
}

export interface TableFooterRow<Row = Record<string, unknown>> extends StyleReference {
  height?: number;
  cells: readonly TableSectionCell<Row>[];
}

export type TableSectionCellAccessor<Row = Record<string, unknown>> = (
  context: TableSectionCellContext<Row>,
) => CellContent;

export interface TableSectionCellContext<Row = Record<string, unknown>> {
  rows: Row[];
  allRows: Row[];
  dataIndex: number;
  rowIndex: number;
}

// ─── Writer types (used by SheetWriter / compileBlock) ────────────────────────

/** Cell data passed to SheetWriter.addCell() */
export interface WriterCell {
  row: number;
  column: number;
  value?: CellValue;
  formula?: string;
  formulaResult?: CellValue;
  link?: WriterLink;
  style?: StyleValue;
  inlineStyle?: CellStyleDefinition;
}

export interface WriterLink {
  target: string;
  tooltip?: string;
}

export interface WriterMergeRange {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

export interface WriterColumnWidth {
  column: number;
  width: number;
}

export interface WriterColumnVisibility {
  column: number;
  hidden: boolean;
}

export interface WriterRowHeight {
  row: number;
  height: number;
}

export interface WriterRowVisibility {
  row: number;
  hidden: boolean;
}

export interface WriterSheetView {
  state: 'frozen';
  xSplit?: number;
  ySplit?: number;
}

// ─── Table column ─────────────────────────────────────────────────────────────

export interface TableColumn<Row = Record<string, unknown>> extends StyleReference {
  title: string;
  id?: keyof Row;
  accessor?: (row: Row) => CellContent;
  children?: readonly TableColumn<Row>[];
  childrenRowOffset?: number;
  width?: number;
  hidden?: boolean;
  headerStyle?: StyleValue;
  bodyStyle?: StyleValue;
  styleResolver?: (value: CellValue | undefined, row: Row, rowIndex: number) => StyleValue | undefined;
  summary?: 'sum' | 'count' | 'average' | FormulaDefinition;
}
