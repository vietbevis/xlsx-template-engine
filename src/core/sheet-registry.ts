import { ReportEngineError } from './errors';
import type { SheetDefinition } from './types';

export type SheetRegistry = ReadonlyMap<string, SheetDefinition>;

export function createSheetRegistry(sheets: readonly SheetDefinition[]): SheetRegistry {
  const registry = new Map<string, SheetDefinition>();

  for (const sheet of sheets) {
    if (registry.has(sheet.id)) {
      throw new ReportEngineError(`Duplicate sheet id "${sheet.id}".`);
    }

    registry.set(sheet.id, sheet);
  }

  return registry;
}
