import { validateWorkbookDefinition } from "./validation";
import type { WorkbookDefinition, WorkbookMetadata } from "./types";

export function defineWorkbook(definition: WorkbookDefinition): WorkbookDefinition {
  validateWorkbookDefinition(definition);

  return {
    ...definition,
    metadata: definition.metadata ? cloneWorkbookMetadata(definition.metadata) : undefined,
    styles: definition.styles ? { ...definition.styles } : undefined,
    sheets: definition.sheets.map((sheet) => ({
      ...sheet,
      blocks: [...sheet.blocks],
    })),
  };
}

export function isWorkbookDefinition(value: unknown): value is WorkbookDefinition {
  try {
    validateWorkbookDefinition(value as WorkbookDefinition);
    return true;
  } catch {
    return false;
  }
}

function cloneWorkbookMetadata(metadata: WorkbookMetadata): WorkbookMetadata {
  return {
    ...metadata,
    keywords: metadata.keywords ? [...metadata.keywords] : undefined,
  };
}
