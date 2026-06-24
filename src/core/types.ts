import type ExcelJS from 'exceljs';

export type CellValue = Exclude<ExcelJS.CellValue, undefined>;

export type CellContent = CellValue | FormulaDefinition;

/** Column node lá (không có children) trong cây cột của table. */
export type TableColumnNode = Extract<Block, { type: 'table' }>['columns'][number];
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
  | NamedRangeFormulaDefinition
  | MaxFormulaDefinition
  | MinFormulaDefinition
  | AverageFormulaDefinition
  | CountFormulaDefinition
  | CountAFormulaDefinition
  | ConcatenateFormulaDefinition
  | IfErrorFormulaDefinition
  | VlookupFormulaDefinition;

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

export interface VlookupFormulaDefinition {
  type: 'vlookup';
  lookup: FormulaDefinition;
  rangeName: string;
  colIndex: number;
  exactMatch?: boolean;
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

export interface NamedRangeFormulaDefinition {
  type: 'namedRange';
  name: string;
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
  namedRanges?: readonly NamedRangeDefinition[];
  sheets: readonly SheetDefinition[];
}

export interface NamedRangeDefinition {
  name: string;
  sheetId: string;
  startId: string;
  endId: string;
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

export type Block = TitleBlock | TextBlock | SpacerBlock | GridBlock | TableBlock | DividerBlock;

export interface BaseBlock {
  type: string;
  context?: Record<string, unknown>;
}

export interface TitleBlock extends BaseBlock, StyleReference {
  type: 'title';
  text: string;
  height?: number;
}

export interface TextBlock extends BaseBlock, StyleReference {
  type: 'text';
  text: string;
  height?: number;
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
  colSpan?: number;
  rowSpan?: number;
  width?: number;
}

export interface TableBlock<Row = Record<string, unknown>> extends BaseBlock {
  type: 'table';
  columns: readonly TableColumn<Row>[];
  data: readonly TableDataItem<Row>[];
  titleRows?: readonly TableTitleRow[];
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

export type TableBorderDefinition = ExcelJS.BorderStyle | Partial<ExcelJS.Borders>;

export type TableDataItem<Row = Record<string, unknown>> = Row | TableSectionRow<Row>;

export interface TableTitleRow extends StyleReference {
  value: CellContent;
  height?: number;
}

export interface TableSectionRow<Row = Record<string, unknown>> extends StyleReference {
  type: 'section';
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

export type TypedWorkbookDefinition<TWorkbook extends WorkbookDefinition> = Omit<
  TWorkbook,
  'namedRanges' | 'sheets'
> & {
  namedRanges?: TypedNamedRanges<TWorkbook>;
  sheets: {
    [TIndex in keyof TWorkbook['sheets']]: TWorkbook['sheets'][TIndex] extends SheetDefinition
      ? TypedSheetDefinition<TWorkbook, TWorkbook['sheets'][TIndex]>
      : TWorkbook['sheets'][TIndex];
  };
};

type TypedSheetDefinition<TWorkbook extends WorkbookDefinition, TSheet extends SheetDefinition> = Omit<
  TSheet,
  'blocks'
> & {
  blocks: {
    [TIndex in keyof TSheet['blocks']]: TSheet['blocks'][TIndex] extends Block
      ? TypedBlockDefinition<TWorkbook, TSheet, TSheet['blocks'][TIndex]>
      : TSheet['blocks'][TIndex];
  };
};

type TypedBlockDefinition<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TBlock extends Block,
> = TBlock extends GridBlock
  ? TypedGridBlock<TWorkbook, TSheet, TBlock>
  : TBlock extends TableBlock<infer TRow>
    ? TypedTableBlock<TWorkbook, TSheet, TBlock, TRow>
    : TBlock extends TitleBlock | TextBlock | DividerBlock
      ? TypedStyleReference<TWorkbook, TBlock>
      : TBlock;

type TypedNamedRanges<TWorkbook extends WorkbookDefinition> = TWorkbook['namedRanges'] extends
  | readonly NamedRangeDefinition[]
  | undefined
  ? {
      [TIndex in keyof TWorkbook['namedRanges']]: TWorkbook['namedRanges'][TIndex] extends NamedRangeDefinition
        ? TypedNamedRangeDefinition<TWorkbook, TWorkbook['namedRanges'][TIndex]>
        : TWorkbook['namedRanges'][TIndex];
    }
  : TWorkbook['namedRanges'];

type TypedNamedRangeDefinition<TWorkbook extends WorkbookDefinition, TRange extends NamedRangeDefinition> = Omit<
  TRange,
  'sheetId' | 'startId' | 'endId'
> &
  {
    [TSheetId in SheetIds<TWorkbook>]: {
      sheetId: TSheetId;
      startId: SheetGridKeys<TWorkbook, TSheetId>;
      endId: SheetGridKeys<TWorkbook, TSheetId>;
    };
  }[SheetIds<TWorkbook>];

type TypedStyleReference<TWorkbook extends WorkbookDefinition, TValue extends StyleReference> = Omit<
  TValue,
  'style'
> & {
  style?: TypedStyleValue<TWorkbook>;
};

type TypedStyleValue<TWorkbook extends WorkbookDefinition> = StyleNames<TWorkbook> | CellStyleDefinition;

type TypedGridBlock<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TBlock extends GridBlock,
> = Omit<TBlock, 'rows'> & {
  rows: {
    [TRowIndex in keyof TBlock['rows']]: TBlock['rows'][TRowIndex] extends GridRow
      ? Omit<TBlock['rows'][TRowIndex], 'cells'> & {
          cells: {
            [TCellIndex in keyof TBlock['rows'][TRowIndex]['cells']]: TBlock['rows'][TRowIndex]['cells'][TCellIndex] extends GridCell
              ? TypedGridCell<TWorkbook, TSheet, TBlock['rows'][TRowIndex]['cells'][TCellIndex]>
              : TBlock['rows'][TRowIndex]['cells'][TCellIndex];
          };
        }
      : TBlock['rows'][TRowIndex];
  };
};

type TypedGridCell<TWorkbook extends WorkbookDefinition, TSheet extends SheetDefinition, TCell extends GridCell> = Omit<
  TCell,
  'value' | 'style' | 'styleResolver'
> & {
  value?: TypedCellContent<TWorkbook, SheetIdOf<TSheet>, SheetGridKeys<TWorkbook, SheetIdOf<TSheet>>>;
  style?: TypedStyleValue<TWorkbook>;
  styleResolver?: (value: CellValue | undefined) => TypedStyleValue<TWorkbook> | undefined;
};

type TypedTableBlock<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TBlock extends TableBlock<TRow>,
  TRow,
> = Omit<TBlock, 'columns'> & {
  headerStyle?: TypedStyleValue<TWorkbook>;
  bodyStyle?: TypedStyleValue<TWorkbook>;
  evenRowStyle?: TypedStyleValue<TWorkbook>;
  oddRowStyle?: TypedStyleValue<TWorkbook>;
  summaryStyle?: TypedStyleValue<TWorkbook>;
  titleRows?: TypedTableTitleRows<TWorkbook, TBlock['titleRows']>;
  footerRows?: TypedTableFooterRows<
    TWorkbook,
    TSheet,
    TRow,
    TBlock['footerRows'],
    TableColumnKeys<TRow, TBlock['columns']>
  >;
  columns: TypedTableColumns<TWorkbook, TSheet, TRow, TBlock['columns'], TableColumnKeys<TRow, TBlock['columns']>>;
};

type TypedTableTitleRows<TWorkbook extends WorkbookDefinition, TRows> = TRows extends readonly TableTitleRow[]
  ? {
      [TIndex in keyof TRows]: TRows[TIndex] extends TableTitleRow
        ? TypedStyleReference<TWorkbook, TRows[TIndex]>
        : TRows[TIndex];
    }
  : TRows;

type TypedTableFooterRows<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TRow,
  TRows,
  TAllKeys extends string,
> = TRows extends readonly TableFooterRow<TRow>[]
  ? {
      [TIndex in keyof TRows]: TRows[TIndex] extends TableFooterRow<TRow>
        ? Omit<TRows[TIndex], 'cells' | 'style'> & {
            style?: TypedStyleValue<TWorkbook>;
            cells: TypedTableSectionCells<TWorkbook, TSheet, TRow, TRows[TIndex]['cells'], TAllKeys>;
          }
        : TRows[TIndex];
    }
  : TRows;

type TypedTableSectionCells<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TRow,
  TCells extends readonly TableSectionCell<TRow>[],
  TAllKeys extends string,
> = {
  [TIndex in keyof TCells]: TCells[TIndex] extends TableSectionCell<TRow>
    ? TypedTableSectionCell<TWorkbook, TSheet, TRow, TCells[TIndex], TAllKeys>
    : TCells[TIndex];
};

type TypedTableSectionCell<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TRow,
  TCell extends TableSectionCell<TRow>,
  TAllKeys extends string,
> = Omit<TCell, 'columnId' | 'value' | 'style' | 'styleResolver'> & {
  columnId?: TAllKeys;
  value?:
    | TypedCellContent<TWorkbook, SheetIdOf<TSheet>, TAllKeys>
    | ((context: TableSectionCellContext<TRow>) => TypedCellContent<TWorkbook, SheetIdOf<TSheet>, TAllKeys>);
  style?: TypedStyleValue<TWorkbook>;
  styleResolver?: (value: CellValue | undefined) => TypedStyleValue<TWorkbook> | undefined;
};

type TypedTableColumns<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TRow,
  TColumns extends readonly TableColumn<TRow>[],
  TAllKeys extends string,
> = {
  [TIndex in keyof TColumns]: TColumns[TIndex] extends TableColumn<TRow>
    ? TypedTableColumn<TWorkbook, TSheet, TRow, TColumns[TIndex], TAllKeys>
    : TColumns[TIndex];
};

type TypedTableColumn<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TRow,
  TColumn extends TableColumn<TRow>,
  TAllKeys extends string,
> = TColumn extends { children: readonly TableColumn<TRow>[] }
  ? Omit<TColumn, 'children' | 'style' | 'headerStyle' | 'bodyStyle' | 'styleResolver'> & {
      style?: TypedStyleValue<TWorkbook>;
      headerStyle?: TypedStyleValue<TWorkbook>;
      bodyStyle?: TypedStyleValue<TWorkbook>;
      styleResolver?: (
        value: CellValue | undefined,
        row: TRow,
        rowIndex: number,
      ) => TypedStyleValue<TWorkbook> | undefined;
      children: TypedTableColumns<TWorkbook, TSheet, TRow, TColumn['children'], TAllKeys>;
    }
  : Omit<TColumn, 'accessor' | 'id' | 'style' | 'headerStyle' | 'bodyStyle' | 'styleResolver' | 'summary'> & {
      id?: Extract<keyof TRow, string>;
      accessor?: (row: TRow) => TypedCellContent<TWorkbook, SheetIdOf<TSheet>, TAllKeys>;
      style?: TypedStyleValue<TWorkbook>;
      headerStyle?: TypedStyleValue<TWorkbook>;
      bodyStyle?: TypedStyleValue<TWorkbook>;
      styleResolver?: (
        value: CellValue | undefined,
        row: TRow,
        rowIndex: number,
      ) => TypedStyleValue<TWorkbook> | undefined;
      summary?: 'sum' | 'count' | 'average' | TypedFormulaDefinition<TWorkbook, SheetIdOf<TSheet>, TAllKeys>;
    };

type TypedCellContent<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = CellValue | TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;

export type TypedFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string = SheetIds<TWorkbook>,
  TLocalKeys extends string = SheetGridKeys<TWorkbook, TCurrentSheetId>,
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
  | NamedRangeFormulaDefinition
  | TypedMaxFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedMinFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedAverageFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedCountFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedCountAFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedConcatenateFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedIfErrorFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>
  | TypedVlookupFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;

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
        id: SheetGridKeys<TWorkbook, TSheetId>;
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
        startId: SheetGridKeys<TWorkbook, TSheetId>;
        endId: SheetGridKeys<TWorkbook, TSheetId>;
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
        startId: SheetGridKeys<TWorkbook, TSheetId>;
        endId: SheetGridKeys<TWorkbook, TSheetId>;
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

type TypedVlookupFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> = {
  type: 'vlookup';
  lookup: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;
  rangeName: string;
  colIndex: number;
  exactMatch?: boolean;
};

type SheetIds<TWorkbook extends WorkbookDefinition> = Extract<TWorkbook['sheets'][number]['id'], string>;

type StyleNames<TWorkbook extends WorkbookDefinition> = Extract<keyof NonNullable<TWorkbook['styles']>, string>;

type SheetIdOf<TSheet extends SheetDefinition> = Extract<TSheet['id'], string>;

type SheetById<TWorkbook extends WorkbookDefinition, TSheetId extends string> = Extract<
  TWorkbook['sheets'][number],
  { id: TSheetId }
>;

type SheetGridKeys<TWorkbook extends WorkbookDefinition, TSheetId extends string> =
  SheetById<TWorkbook, TSheetId> extends infer TSheet
    ? TSheet extends SheetDefinition
      ? TSheet['blocks'][number] extends infer TBlock
        ? TBlock extends GridBlock
          ? GridBlockKeys<TBlock>
          : never
        : never
      : never
    : never;

type GridBlockKeys<TBlock extends GridBlock> = TBlock['rows'][number]['cells'][number] extends infer TCell
  ? TCell extends { id: infer TKey }
    ? Extract<TKey, string>
    : never
  : never;

type TableColumnKeys<TRow, TColumns extends readonly TableColumn<TRow>[]> = TColumns[number] extends infer TColumn
  ? TColumn extends TableColumn<TRow>
    ? TColumn extends { children: readonly TableColumn<TRow>[] }
      ? TableColumnKeys<TRow, TColumn['children']>
      : TColumn extends { id: infer TKey }
        ? Extract<TKey, string>
        : never
    : never
  : never;
