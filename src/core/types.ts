export type CellValue = string | number | boolean | Date | null;

export type StyleRegistry = Record<string, CellStyleDefinition>;

export interface StyleReference {
  style?: string;
}

export interface CellStyleDefinition {
  font?: FontStyleDefinition;
  fill?: FillStyleDefinition;
  border?: BorderStyleDefinition;
  alignment?: AlignmentStyleDefinition;
  numberFormat?: string;
}

export interface FontStyleDefinition {
  name?: string;
  size?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  color?: ColorStyleDefinition;
}

export interface FillStyleDefinition {
  pattern?: FillPatternStyle;
  foregroundColor?: ColorStyleDefinition;
  backgroundColor?: ColorStyleDefinition;
}

export interface BorderStyleDefinition {
  top?: BorderSideStyleDefinition;
  right?: BorderSideStyleDefinition;
  bottom?: BorderSideStyleDefinition;
  left?: BorderSideStyleDefinition;
}

export interface BorderSideStyleDefinition {
  style: BorderLineStyle;
  color?: ColorStyleDefinition;
}

export interface AlignmentStyleDefinition {
  horizontal?: HorizontalAlignmentStyle;
  vertical?: VerticalAlignmentStyle;
  wrapText?: boolean;
}

export interface ColorStyleDefinition {
  argb: string;
}

export type FillPatternStyle =
  | "none"
  | "solid"
  | "darkVertical"
  | "darkHorizontal"
  | "darkGrid"
  | "darkTrellis"
  | "darkDown"
  | "darkUp"
  | "lightVertical"
  | "lightHorizontal"
  | "lightGrid"
  | "lightTrellis"
  | "lightDown"
  | "lightUp"
  | "darkGray"
  | "mediumGray"
  | "lightGray"
  | "gray125"
  | "gray0625";

export type BorderLineStyle =
  | "thin"
  | "dotted"
  | "hair"
  | "medium"
  | "double"
  | "thick"
  | "dashDot"
  | "dashDotDot"
  | "slantDashDot"
  | "mediumDashed"
  | "mediumDashDotDot"
  | "mediumDashDot";

export type HorizontalAlignmentStyle =
  | "left"
  | "center"
  | "right"
  | "fill"
  | "justify"
  | "centerContinuous"
  | "distributed";

export type VerticalAlignmentStyle =
  | "top"
  | "middle"
  | "bottom"
  | "distributed"
  | "justify";

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
  sheets: SheetDefinition[];
}

export interface SheetDefinition {
  id: string;
  name: string;
  blocks: Block[];
}

export type Block =
  | TitleBlock
  | TextBlock
  | SpacerBlock
  | GridBlock
  | TableBlock<any>;

export interface BaseBlock {
  type: string;
}

export interface TitleBlock extends BaseBlock, StyleReference {
  type: "title";
  text: string;
  height?: number;
}

export interface TextBlock extends BaseBlock, StyleReference {
  type: "text";
  text: string;
  height?: number;
}

export interface SpacerBlock extends BaseBlock {
  type: "spacer";
  rows?: number;
}

export interface GridBlock extends BaseBlock {
  type: "grid";
  rows: GridRow[];
}

export interface GridRow {
  height?: number;
  cells: GridCell[];
}

export interface GridCell extends StyleReference {
  value?: CellValue;
  colSpan?: number;
  rowSpan?: number;
  width?: number;
}

export interface TableBlock<Row = Record<string, unknown>> extends BaseBlock {
  type: "table";
  columns: TableColumn<Row>[];
  data: Row[] | AsyncIterable<Row>;
  headerStyle?: string;
  bodyStyle?: string;
}

export interface TableColumn<Row = Record<string, unknown>> extends StyleReference {
  title: string;
  key?: string;
  accessor?: (row: Row) => CellValue;
  children?: TableColumn<Row>[];
  width?: number;
}
