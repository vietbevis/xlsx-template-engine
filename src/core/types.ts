export type CellValue = string | number | boolean | Date | null;

export type CellContent = CellValue | FormulaDefinition;

export type FormulaDefinition =
  | RawFormulaDefinition
  | LiteralFormulaDefinition
  | SumFormulaDefinition
  | RoundFormulaDefinition
  | IfFormulaDefinition
  | CallFormulaDefinition
  | BinaryFormulaDefinition
  | RangeFormulaDefinition
  | RefFormulaDefinition;

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
  values?: FormulaDefinition[];
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
  args: FormulaDefinition[];
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
  key: string;
}

export interface RangeFormulaDefinition {
  type: 'range';
  sheetId?: string;
  startKey: string;
  endKey: string;
  scope?: FormulaRangeScope;
}

export interface FormulaRangeReference {
  sheetId?: string;
  startKey: string;
  endKey: string;
  scope?: FormulaRangeScope;
}

export type FormulaRangeScope = 'currentRows' | 'allRows';

export type FormulaBinaryOperator = '+' | '-' | '*' | '/' | '>' | '>=' | '<' | '<=' | '=' | '<>';

export type StyleRegistry = Record<string, CellStyleDefinition>;

export type StyleValue = string | CellStyleDefinition;

export interface StyleReference {
  style?: StyleValue;
}

export interface CellStyleDefinition {
  extends?: string;
  [key: string]: unknown;
  font?: FontStyleDefinition;
  fill?: FillStyleDefinition;
  border?: BorderStyleDefinition;
  alignment?: AlignmentStyleDefinition;
  numFmt?: string;
  numberFormat?: string;
  protection?: Record<string, unknown>;
}

export interface FontStyleDefinition {
  [key: string]: unknown;
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: ColorStyleDefinition;
}

export interface FillStyleDefinition {
  [key: string]: unknown;
  type?: string;
  pattern?: FillPatternStyle;
  fgColor?: ColorStyleDefinition;
  bgColor?: ColorStyleDefinition;
  foregroundColor?: ColorStyleDefinition;
  backgroundColor?: ColorStyleDefinition;
}

export interface BorderStyleDefinition {
  [key: string]: unknown;
  top?: BorderSideStyleDefinition;
  right?: BorderSideStyleDefinition;
  bottom?: BorderSideStyleDefinition;
  left?: BorderSideStyleDefinition;
}

export interface BorderSideStyleDefinition {
  [key: string]: unknown;
  style: BorderLineStyle;
  color?: ColorStyleDefinition;
}

export interface AlignmentStyleDefinition {
  [key: string]: unknown;
  horizontal?: HorizontalAlignmentStyle;
  vertical?: VerticalAlignmentStyle;
  wrapText?: boolean;
}

export interface ColorStyleDefinition {
  [key: string]: unknown;
  argb: string;
}

export type FillPatternStyle =
  | 'none'
  | 'solid'
  | 'darkVertical'
  | 'darkHorizontal'
  | 'darkGrid'
  | 'darkTrellis'
  | 'darkDown'
  | 'darkUp'
  | 'lightVertical'
  | 'lightHorizontal'
  | 'lightGrid'
  | 'lightTrellis'
  | 'lightDown'
  | 'lightUp'
  | 'darkGray'
  | 'mediumGray'
  | 'lightGray'
  | 'gray125'
  | 'gray0625';

export type BorderLineStyle =
  | 'thin'
  | 'dotted'
  | 'hair'
  | 'medium'
  | 'double'
  | 'thick'
  | 'dashDot'
  | 'dashDotDot'
  | 'slantDashDot'
  | 'mediumDashed'
  | 'mediumDashDotDot'
  | 'mediumDashDot';

export type HorizontalAlignmentStyle =
  | 'left'
  | 'center'
  | 'right'
  | 'fill'
  | 'justify'
  | 'centerContinuous'
  | 'distributed';

export type VerticalAlignmentStyle = 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';

export interface WorkbookMetadata {
  title?: string;
  author?: string;
  company?: string;
  subject?: string;
  keywords?: string[];
}

export interface WorkbookDefinition {
  metadata?: WorkbookMetadata;
  styles?: StyleRegistry;
  context?: Record<string, unknown>;
  sheets: SheetDefinition[];
}

export interface SheetDefinition {
  id: string;
  name: string;
  context?: Record<string, unknown>;
  blocks: Block[];
}

export type Block = TitleBlock | TextBlock | SpacerBlock | GridBlock | TableBlock<any>;

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

export interface GridBlock extends BaseBlock {
  type: 'grid';
  rows: GridRow[];
}

export interface GridRow {
  height?: number;
  cells: GridCell[];
}

export interface GridCell extends StyleReference {
  key?: string;
  value?: CellContent;
  colSpan?: number;
  rowSpan?: number;
  width?: number;
}

export interface TableBlock<Row = Record<string, unknown>> extends BaseBlock {
  type: 'table';
  columns: TableColumn<Row>[];
  data: TableDataItem<Row>[] | AsyncIterable<TableDataItem<Row>>;
  titleRows?: TableTitleRow[];
  headerStyle?: StyleValue;
  bodyStyle?: StyleValue;
  border?: TableBorderDefinition;
}

export type TableBorderDefinition = BorderLineStyle | BorderStyleDefinition;

export type TableDataItem<Row = Record<string, unknown>> = Row | TableSectionRow<Row>;

export interface TableTitleRow extends StyleReference {
  value: CellContent;
  height?: number;
}

export interface TableSectionRow<Row = Record<string, unknown>> extends StyleReference {
  type: 'section';
  resetRows?: boolean;
  height?: number;
  cells: TableSectionCell<Row>[];
}

export interface TableSectionCell<Row = Record<string, unknown>> extends StyleReference {
  key?: string;
  column?: number;
  columnKey?: string;
  value?: CellContent | TableSectionCellAccessor<Row>;
  colSpan?: number | 'remaining';
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
  key?: keyof Row;
  accessor?: (row: Row) => CellContent;
  children?: TableColumn<Row>[];
  childrenRowOffset?: number;
  width?: number;
  headerStyle?: StyleValue;
  bodyStyle?: StyleValue;
}

export type TypedWorkbookDefinition<TWorkbook extends WorkbookDefinition> = Omit<
  TWorkbook,
  'sheets'
> & {
  sheets: {
    [TIndex in keyof TWorkbook['sheets']]: TWorkbook['sheets'][TIndex] extends SheetDefinition
      ? TypedSheetDefinition<TWorkbook, TWorkbook['sheets'][TIndex]>
      : TWorkbook['sheets'][TIndex];
  };
};

type TypedSheetDefinition<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
> = Omit<TSheet, 'blocks'> & {
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
    : TBlock;

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

type TypedGridCell<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TCell extends GridCell,
> = Omit<TCell, 'value'> & {
  value?: TypedCellContent<
    TWorkbook,
    SheetIdOf<TSheet>,
    SheetGridKeys<TWorkbook, SheetIdOf<TSheet>>
  >;
};

type TypedTableBlock<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TBlock extends TableBlock<TRow>,
  TRow,
> = Omit<TBlock, 'columns'> & {
  columns: TypedTableColumns<
    TWorkbook,
    TSheet,
    TRow,
    TBlock['columns'],
    TableColumnKeys<TBlock['columns']>
  >;
};

type TypedTableColumns<
  TWorkbook extends WorkbookDefinition,
  TSheet extends SheetDefinition,
  TRow,
  TColumns extends TableColumn<TRow>[],
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
> = TColumn extends { children: TableColumn<TRow>[] }
  ? Omit<TColumn, 'children'> & {
      children: TypedTableColumns<TWorkbook, TSheet, TRow, TColumn['children'], TAllKeys>;
    }
  : Omit<TColumn, 'accessor'> & {
      accessor?: (row: TRow) => TypedCellContent<TWorkbook, SheetIdOf<TSheet>, TAllKeys>;
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
  | TypedBinaryFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>;

type TypedRefFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> =
  | { type: 'ref'; key: TLocalKeys; sheetId?: undefined }
  | {
      [TSheetId in SheetIds<TWorkbook>]: {
        type: 'ref';
        sheetId: TSheetId;
        key: SheetGridKeys<TWorkbook, TSheetId>;
      };
    }[SheetIds<TWorkbook>];

type TypedRangeFormulaDefinition<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> =
  | {
      type: 'range';
      startKey: TLocalKeys;
      endKey: TLocalKeys;
      scope?: FormulaRangeScope;
      sheetId?: undefined;
    }
  | {
      [TSheetId in SheetIds<TWorkbook>]: {
        type: 'range';
        sheetId: TSheetId;
        startKey: SheetGridKeys<TWorkbook, TSheetId>;
        endKey: SheetGridKeys<TWorkbook, TSheetId>;
        scope?: undefined;
      };
    }[SheetIds<TWorkbook>];

type TypedFormulaRangeReference<
  TWorkbook extends WorkbookDefinition,
  TCurrentSheetId extends string,
  TLocalKeys extends string,
> =
  | {
      startKey: TLocalKeys;
      endKey: TLocalKeys;
      scope?: FormulaRangeScope;
      sheetId?: undefined;
    }
  | {
      [TSheetId in SheetIds<TWorkbook>]: {
        sheetId: TSheetId;
        startKey: SheetGridKeys<TWorkbook, TSheetId>;
        endKey: SheetGridKeys<TWorkbook, TSheetId>;
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
  values?: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>[];
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
  args: TypedFormulaDefinition<TWorkbook, TCurrentSheetId, TLocalKeys>[];
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

type SheetIds<TWorkbook extends WorkbookDefinition> = Extract<
  TWorkbook['sheets'][number]['id'],
  string
>;

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

type GridBlockKeys<TBlock extends GridBlock> =
  TBlock['rows'][number]['cells'][number] extends infer TCell
    ? TCell extends { key: infer TKey }
      ? Extract<TKey, string>
      : never
    : never;

type TableColumnKeys<TColumns extends TableColumn<any>[]> = TColumns[number] extends infer TColumn
  ? TColumn extends TableColumn<any>
    ? TColumn extends { children: TableColumn<any>[] }
      ? TableColumnKeys<TColumn['children']>
      : TColumn extends { key: infer TKey }
        ? Extract<TKey, string>
        : never
    : never
  : never;
