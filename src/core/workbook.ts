import { validateWorkbookDefinition } from './validation';
import type { TypedWorkbookDefinition, WorkbookDefinition, WorkbookMetadata } from './types';

export function defineWorkbook<const TWorkbook extends WorkbookDefinition>(
  definition: TWorkbook & TypedWorkbookDefinition<TWorkbook>,
): TWorkbook {
  validateWorkbookDefinition(definition);

  return {
    ...definition,
    metadata: definition.metadata ? cloneWorkbookMetadata(definition.metadata) : undefined,
    styles: definition.styles ? { ...definition.styles } : undefined,
    context: definition.context ? { ...definition.context } : undefined,
    sheets: definition.sheets.map((sheet) => ({
      ...sheet,
      context: sheet.context ? { ...sheet.context } : undefined,
      blocks: [...sheet.blocks],
    })),
  } as TWorkbook;
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
