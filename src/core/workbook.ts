import { validateWorkbookDefinition } from './validation';
import type {
  CellStyleDefinition,
  TypedWorkbookDefinition,
  WorkbookDefinition,
  WorkbookMetadata,
} from './types';

export function defineWorkbook<const TWorkbook extends WorkbookDefinition>(
  definition: TWorkbook & TypedWorkbookDefinition<TWorkbook>,
): TWorkbook {
  validateWorkbookDefinition(definition);

  return {
    ...definition,
    metadata: definition.metadata ? cloneWorkbookMetadata(definition.metadata) : undefined,
    defaultStyle: definition.defaultStyle ? clonePlainObject(definition.defaultStyle) : undefined,
    styles: definition.styles ? { ...definition.styles } : undefined,
    context: definition.context ? { ...definition.context } : undefined,
    sheets: definition.sheets.map((sheet) => ({
      ...sheet,
      context: sheet.context ? { ...sheet.context } : undefined,
      blocks: [...sheet.blocks],
    })),
  } as TWorkbook;
}

function clonePlainObject<T extends CellStyleDefinition>(value: T): T {
  return clonePart(value) as T;
}

function clonePart(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((item) => clonePart(item));
  }

  if (typeof value !== 'object' || value === null) {
    return value;
  }

  return Object.fromEntries(Object.entries(value).map(([key, child]) => [key, clonePart(child)]));
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
