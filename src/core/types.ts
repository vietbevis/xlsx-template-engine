export type CellValue = string | number | boolean | Date | null;

export type StyleRegistry = Record<string, unknown>;

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
  | TableBlock;

export interface BaseBlock {
  type: string;
}

export interface TitleBlock extends BaseBlock {
  type: "title";
  text: CellValue;
  style?: string;
}

export interface TextBlock extends BaseBlock {
  type: "text";
  text: CellValue;
  style?: string;
}

export interface SpacerBlock extends BaseBlock {
  type: "spacer";
  rows?: number;
}

export interface GridBlock extends BaseBlock {
  type: "grid";
}

export interface TableBlock extends BaseBlock {
  type: "table";
}
