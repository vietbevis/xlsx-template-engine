import { ValidationError } from './errors';
import { isPlainObject } from './helpers/common';
import { calculateTableHeaderDepth } from './helpers/table';
import type { WorkbookDefinition } from './types';

// ─── Public entry point ───────────────────────────────────────────────────────

export function validateWorkbookDefinition(workbook: WorkbookDefinition): void {
  if (!workbook || !Array.isArray(workbook.sheets) || workbook.sheets.length === 0) {
    throw new ValidationError('Workbook definition must include at least one sheet.');
  }

  validateSemantics(workbook);
}

// ─── Semantic validation ──────────────────────────────────────────────────────

function validateSemantics(workbook: WorkbookDefinition): void {
  validateStyleRegistryNames(workbook.styles);

  const seenIds = new Set<string>();
  const seenNames = new Set<string>();

  for (const sheet of workbook.sheets) {
    validateSheetName(sheet.name, sheet.id);

    if (seenIds.has(sheet.id)) throw new ValidationError(`Duplicate sheet id "${sheet.id}".`);
    if (seenNames.has(sheet.name.trim().toLowerCase()))
      throw new ValidationError(`Duplicate sheet name "${sheet.name}".`);

    seenIds.add(sheet.id);
    seenNames.add(sheet.name.trim().toLowerCase());

    for (const [blockIndex, block] of sheet.blocks.entries()) {
      const blockLabel = blockPath(sheet.id, block, blockIndex);
      validateBlockSemantics(block, blockLabel, workbook.styles);
    }
  }
}

function validateBlockSemantics(
  block: WorkbookDefinition['sheets'][number]['blocks'][number],
  label: string,
  styles: WorkbookDefinition['styles'],
): void {
  validateStyleRefs(block, label, styles);

  if (block.type !== 'table') return;

  if (block.headerRowHeights && block.headerRowHeights.length > calculateTableHeaderDepth(block.columns)) {
    throw new ValidationError(`${label} headerRowHeights must not exceed header row count.`);
  }

  const rows = block.data ?? block.groups!.flatMap((g) => g.data);
  validateTableDataRows(rows, `${label} data`);
}

// ─── Style reference validation ───────────────────────────────────────────────

const STYLE_KEYS = ['style', 'headerStyle', 'bodyStyle', 'evenRowStyle', 'oddRowStyle', 'summaryStyle'] as const;

function validateStyleRefs(value: unknown, label: string, styles: WorkbookDefinition['styles']): void {
  if (!isPlainObject(value)) return;

  for (const key of STYLE_KEYS) {
    const style = value[key];
    if (typeof style === 'string' && !styles?.[style]) {
      throw new ValidationError(`${label} ${key} references unknown style "${style}".`);
    }
  }

  for (const [key, child] of Object.entries(value)) {
    if (key === 'styles') continue;

    if (Array.isArray(child)) {
      child.forEach((item, i) => validateStyleRefs(item, `${label} ${key} ${i}`, styles));
    } else if (isPlainObject(child) && typeof child.type !== 'string') {
      validateStyleRefs(child, `${label} ${key}`, styles);
    }
  }
}

function validateStyleRegistryNames(styles: WorkbookDefinition['styles']): void {
  for (const name of Object.keys(styles ?? {})) {
    if (name.trim() === '') throw new ValidationError('Workbook style names must be non-empty.');
  }
}

// ─── Table data row validation ────────────────────────────────────────────────

function validateTableDataRows(rows: readonly unknown[], label: string): void {
  for (const [i, item] of rows.entries()) {
    if (isPlainObject(item) && item.type === 'section') {
      throw new ValidationError(`${label} ${i} is a section row. Move it to headerRows or footerRows.`);
    }
  }
}

// ─── Sheet name validation ────────────────────────────────────────────────────

const INVALID_SHEET_CHARS = /[:\\/?*[\]]/;

function validateSheetName(name: string, id: string): void {
  if (name.length > 31) {
    throw new ValidationError(`Sheet "${id}" name must be 31 characters or fewer.`);
  }
  if (INVALID_SHEET_CHARS.test(name)) {
    throw new ValidationError(`Sheet "${id}" name contains characters Excel does not allow.`);
  }
}

function blockPath(
  sheetId: string,
  block: WorkbookDefinition['sheets'][number]['blocks'][number],
  index: number,
): string {
  return `sheet "${sheetId}" > ${block.type} block ${index + 1}`;
}
