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

export interface Block {
  type: string;
  [property: string]: unknown;
}
