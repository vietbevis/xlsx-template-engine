import type ExcelJS from 'exceljs';

export type CellValue = Exclude<ExcelJS.CellValue, undefined>;

export type CellContent = CellValue | FormulaDefinition;

/** Column node lá (không có children) trong cây cột của table. */
export type TableColumnNode = Extract<Block, { type: 'table' | 'table-groups' }>['columns'][number];
export type TableLeafColumn = TableColumnNode & { children?: undefined };

export type FormulaDefinition =
  | RawFormulaDefinition
  | LiteralFormulaDefinition
  | SumFormulaDefinition
  | RoundFormulaDefinition
  | IfFormulaDefinition
  | CallFormulaDefinition
  | BinaryFormulaDefinition
  | RangeFormulaDefinition
  | RefFormulaDefinition
  | MaxFormulaDefinition
  | MinFormulaDefinition
  | AverageFormulaDefinition
  | CountFormulaDefinition
  | CountAFormulaDefinition
  | ConcatenateFormulaDefinition
  | IfErrorFormulaDefinition;

export interface RawFormulaDefinition {
  type: 'raw';
  expression: string;
}

export interface LiteralFormulaDefinition {
  type: 'literal';
  value: string | number | boolean | null;
}

export interface SumFormulaDefinition {
  type: 'sum';
  range?: FormulaRangeReference;
  values?: readonly FormulaDefinition[];
}

export interface RoundFormulaDefinition {
  type: 'round';
  value: FormulaDefinition;
  digits: number;
}

export interface IfFormulaDefinition {
  type: 'if';
  condition: FormulaDefinition;
  whenTrue: FormulaDefinition;
  whenFalse: FormulaDefinition;
}

export interface CallFormulaDefinition {
  type: 'call';
  name: string;
  args: readonly FormulaDefinition[];
}

export interface MaxFormulaDefinition {
  type: 'max';
  values: readonly FormulaDefinition[];
}

export interface MinFormulaDefinition {
  type: 'min';
  values: readonly FormulaDefinition[];
}

export interface AverageFormulaDefinition {
  type: 'average';
  range: FormulaRangeReference;
}

export interface CountFormulaDefinition {
  type: 'count';
  range: FormulaRangeReference;
}

export interface CountAFormulaDefinition {
  type: 'counta';
  range: FormulaRangeReference;
}

export interface ConcatenateFormulaDefinition {
  type: 'concatenate';
  values: readonly FormulaDefinition[];
}

export interface IfErrorFormulaDefinition {
  type: 'iferror';
  value: FormulaDefinition;
  fallback: FormulaDefinition;
}

export interface BinaryFormulaDefinition {
  type: 'binary';
  operator: FormulaBinaryOperator;
  left: FormulaDefinition;
  right: FormulaDefinition;
}

export interface RefFormulaDefinition {
  type: 'ref';
  sheetId?: string;
  id: string;
}

export interface RangeFormulaDefinition {
  type: 'range';
  sheetId?: string;
  startId: string;
  endId: string;
  scope?: FormulaRangeScope;
}

export interface FormulaRangeReference {
  sheetId?: string;
  startId: string;
  endId: string;
  scope?: FormulaRangeScope;
}

export type FormulaRangeScope = 'currentRows' | 'allRows';

export type FormulaBinaryOperator = '+' | '-' | '*' | '/' | '>' | '>=' | '<' | '<=' | '=' | '<>';

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
  context?: Record<string, unknown>;
  sheets: readonly SheetDefinition[];
}

export interface SheetDefinition {
  id: string;
  name: string;
  context?: Record<string, unknown>;
  freezePane?: SheetFreezePane;
  blocks: readonly Block[];
}

export interface SheetFreezePane {
  rows?: number;
  columns?: number;
}

export type Block = TitleBlock | TextBlock | SpacerBlock | GridBlock | TableBlock | TableGroupsBlock | DividerBlock;

export interface BaseBlock {
  type: string;
  context?: Record<string, unknown>;
}

export interface TitleBlock extends BaseBlock, StyleReference {
  type: 'title';
  text: string;
  height?: number;
  colSpan?: number | 'remaining';
}

export interface TextBlock extends BaseBlock, StyleReference {
  type: 'text';
  text: string;
  height?: number;
  colSpan?: number | 'remaining';
}

export interface SpacerBlock extends BaseBlock {
  type: 'spacer';
  rows?: number;
}

export interface DividerBlock extends BaseBlock, StyleReference {
  type: 'divider';
  rows?: number;
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
  data: readonly Row[];
}

export interface TableGroupsBlock<Row = Record<string, unknown>> extends BaseTableBlock<Row> {
  type: 'table-groups';
  groups: readonly TableGroup<Row>[];
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

export type TypedFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string = SheetIds<TWorkbook>,
  TLocalKeys extends string = SheetFormulaKeys<TWorkbook, TCurrentSheetId>,
> =
  | RawFormulaDefinition
  | LiteralFormulaDefinition
  | TypedRefFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedRangeFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedSumFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedRoundFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedIfFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedCallFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedBinaryFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedMaxFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedMinFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedAverageFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedCountFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedCountAFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedConcatenateFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedIfErrorFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;

type TypedRefFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> =
  | { type: 'ref'; id: TLocalKeys; sheetId?: undefined }
  | {
      [TSheetId in SheetIds<TWorkbook>]: {
        type: 'ref';
        sheetId: TSheetId;
        id: SheetFormulaKeys<TWorkbook, TSheetId>;
      };
    }[SheetIds<TWorkbook>];

type TypedRangeFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> =
  | {
      type: 'range';
      startId: TLocalKeys;
      endId: TLocalKeys;
      scope?: FormulaRangeScope;
      sheetId?: undefined;
    }
  | {
      [TSheetId in SheetIds<TWorkbook>]: {
        type: 'range';
        sheetId: TSheetId;
        startId: SheetFormulaKeys<TWorkbook, TSheetId>;
        endId: SheetFormulaKeys<TWorkbook, TSheetId>;
        scope?: undefined;
      };
    }[SheetIds<TWorkbook>];

type TypedFormulaRangeReference<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> =
  | {
      startId: TLocalKeys;
      endId: TLocalKeys;
      scope?: FormulaRangeScope;
      sheetId?: undefined;
    }
  | {
      [TSheetId in SheetIds<TWorkbook>]: {
        sheetId: TSheetId;
        startId: SheetFormulaKeys<TWorkbook, TSheetId>;
        endId: SheetFormulaKeys<TWorkbook, TSheetId>;
        scope?: undefined;
      };
    }[SheetIds<TWorkbook>];

type TypedSumFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'sum';
  range?: TypedFormulaRangeReference<TWorkbook, TCurrentSheetId, TLocalKeys>;
  values?: readonly TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>[];
};

type TypedRoundFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'round';
  value: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;
  digits: number;
};

type TypedIfFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'if';
  condition: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;
  whenTrue: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;
  whenFalse: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;
};

type TypedCallFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'call';
  name: string;
  args: readonly TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>[];
};

type TypedBinaryFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'binary';
  operator: FormulaBinaryOperator;
  left: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;
  right: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;
};

type TypedMaxFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'max';
  values: readonly TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>[];
};

type TypedMinFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'min';
  values: readonly TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>[];
};

type TypedAverageFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'average';
  range: TypedFormulaRangeReference<TWorkbook, TCurrentSheetId, TLocalKeys>;
};

type TypedCountFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'count';
  range: TypedFormulaRangeReference<TWorkbook, TCurrentSheetId, TLocalKeys>;
};

type TypedCountAFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'counta';
  range: TypedFormulaRangeReference<TWorkbook, TCurrentSheetId, TLocalKeys>;
};

type TypedConcatenateFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'concatenate';
  values: readonly TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>[];
};

type TypedIfErrorFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'iferror';
  value: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;
  fallback: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;
};

type SheetIds<TWorkbook extends WorkbookDefinition> = Extract<TWorkbook['sheets'][number]['id'], string>;

type SheetById<TWorkbook extends WorkbookDefinition, TSheetId extends string> = Extract<
  TWorkbook['sheets'][number],
  { id: TSheetId }
>;

type SheetFormulaKeys<TWorkbook extends WorkbookDefinition, TSheetId extends string> =
  SheetById<TWorkbook, TSheetId> extends infer TSheet
    ? TSheet extends SheetDefinition
      ? TSheet['blocks'][number] extends infer TBlock
        ? BlockFormulaKeys<TBlock>
        : never
      : never
    : never;

type BlockFormulaKeys<TBlock> =
  | (TBlock extends GridBlock ? GridBlockKeys<TBlock> : never)
  | (TBlock extends TableBlock ? TableBlockKeys<TBlock> : never)
  | (TBlock extends TableGroupsBlock ? TableGroupsBlockKeys<TBlock> : never);

type GridBlockKeys<TBlock extends GridBlock> = TBlock['rows'][number]['cells'][number] extends infer TCell
  ? TCell extends { id: infer TKey }
    ? Extract<TKey, string>
    : never
  : never;

type TableBlockKeys<TBlock extends TableBlock> =
  | TableColumnKeys<TBlock['columns']>
  | TableSectionRowsKeys<TBlock['footerRows']>;

type TableGroupsBlockKeys<TBlock extends TableGroupsBlock> =
  | TableColumnKeys<TBlock['columns']>
  | TableSectionRowsKeys<TBlock['footerRows']>
  | TableGroupKeys<TBlock['groups']>;

type TableGroupKeys<TGroups> = TGroups extends readonly (infer TGroup)[]
  ? TGroup extends TableGroup
    ? TableSectionRowsKeys<TGroup['headerRows']> | TableSectionRowsKeys<TGroup['footerRows']>
    : never
  : never;

type TableSectionRowsKeys<TRows> = TRows extends readonly (infer TRow)[]
  ? TRow extends TableSectionRow
    ? TableSectionCellKeys<TRow['cells']>
    : never
  : never;

type TableSectionCellKeys<TCells> = TCells extends readonly (infer TCell)[]
  ? TCell extends { id: infer TKey }
    ? Extract<TKey, string>
    : never
  : never;

type TableColumnKeys<TColumns> = TColumns extends readonly (infer TColumn)[]
  ? TColumn extends TableColumn
    ? TColumn extends { children: infer TChildren }
      ? TableColumnKeys<TChildren>
      : TColumn extends { id: infer TKey }
        ? Extract<TKey, string>
        : never
    : never
  : never;
