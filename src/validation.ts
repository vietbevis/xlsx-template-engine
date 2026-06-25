import { z, type ZodIssue } from 'zod';
import { ValidationError } from './errors';
import { isPlainObject } from './helpers/common';
import { calculateTableHeaderDepth } from './helpers/table';
import type { WorkbookDefinition } from './types';

// ─── Primitives ───────────────────────────────────────────────────────────────

const nonEmptyString = z.string().trim().min(1);
const positiveNumber = z.number().positive();
const positiveInteger = z.number().int().positive();
const booleanFlag = z.boolean();
const recordSchema = z.record(z.string(), z.unknown());
const styleObjectSchema = z.record(z.string(), z.unknown());
const styleValueSchema = z.union([nonEmptyString, styleObjectSchema]);
const colSpanSchema = z.union([positiveInteger, z.literal('remaining')]);
const functionSchema = z.custom<(...args: never[]) => unknown>((value) => typeof value === 'function', {
  message: 'must be a function',
});

// ─── Formula ──────────────────────────────────────────────────────────────────

const formulaSchema: z.ZodType<unknown> = z.lazy(() =>
  z.discriminatedUnion('type', [
    z.object({
      type: z.literal('formula_template'),
      strings: z.array(z.string()),
      exprs: z.array(z.union([z.string(), z.number(), z.boolean(), z.null(), formulaSchema])),
    }),
    z.object({
      type: z.literal('range'),
      sheetId: nonEmptyString.optional(),
      startId: nonEmptyString,
      endId: nonEmptyString,
      scope: z.enum(['currentRows', 'allRows']).optional(),
    }),
    z.object({
      type: z.literal('ref'),
      sheetId: nonEmptyString.optional(),
      id: nonEmptyString,
    }),
  ]),
);

// Primitive cell values + formula objects
const cellContentSchema = z.union([z.null(), z.string(), z.number(), z.boolean(), z.date(), formulaSchema]);

// ─── Grid ─────────────────────────────────────────────────────────────────────

const gridCellSchema = z.object({
  id: nonEmptyString.optional(),
  value: cellContentSchema.optional(),
  formulaResult: z.unknown().optional(),
  style: styleValueSchema.optional(),
  styleResolver: functionSchema.optional(),
  colSpan: colSpanSchema.optional(),
  rowSpan: positiveInteger.optional(),
  width: positiveNumber.optional(),
});

const gridRowSchema = z.object({
  height: positiveNumber.optional(),
  cells: z.array(gridCellSchema),
});

// ─── Table ────────────────────────────────────────────────────────────────────

const tableSectionCellSchema = z.object({
  id: nonEmptyString.optional(),
  column: positiveInteger.optional(),
  columnId: nonEmptyString.optional(),
  value: z.union([cellContentSchema, functionSchema]).optional(),
  formulaResult: z.unknown().optional(),
  style: styleValueSchema.optional(),
  styleResolver: functionSchema.optional(),
  colSpan: colSpanSchema.optional(),
});

/** Shared shape for header/footer section rows — footer omits resetRows & hidden. */
const baseSectionRowShape = {
  height: positiveNumber.optional(),
  style: styleValueSchema.optional(),
  cells: z.array(tableSectionCellSchema).nonempty(),
} as const;

const tableSectionRowSchema = z.object({
  ...baseSectionRowShape,
  resetRows: booleanFlag.optional(),
  hidden: booleanFlag.optional(),
});

const tableFooterRowSchema = z.object(baseSectionRowShape);

const tableColumnSchema: z.ZodType<unknown> = z.lazy(() =>
  z
    .object({
      title: nonEmptyString,
      id: nonEmptyString.optional(),
      accessor: functionSchema.optional(),
      children: z.array(tableColumnSchema).nonempty().optional(),
      childrenRowOffset: positiveInteger.optional(),
      width: positiveNumber.optional(),
      hidden: booleanFlag.optional(),
      style: styleValueSchema.optional(),
      headerStyle: styleValueSchema.optional(),
      bodyStyle: styleValueSchema.optional(),
      styleResolver: functionSchema.optional(),
      summary: z.union([z.enum(['sum', 'count', 'average']), formulaSchema]).optional(),
    })
    .superRefine((col, ctx) => {
      if (col.children) {
        if (col.id !== undefined || col.accessor !== undefined) {
          ctx.addIssue({ code: 'custom', message: 'with children must not include id or accessor' });
        }
        return;
      }
      if (col.childrenRowOffset !== undefined) {
        ctx.addIssue({ code: 'custom', message: 'childrenRowOffset requires children' });
      }
      if (col.id === undefined && col.accessor === undefined) {
        ctx.addIssue({ code: 'custom', message: 'leaf column must include an id or accessor' });
      }
    }),
);

const tableGroupSchema = z.object({
  headerRows: z.array(tableSectionRowSchema).optional(),
  data: z.array(recordSchema),
  footerRows: z.array(tableSectionRowSchema).optional(),
});

const baseTableShape = {
  columns: z.array(tableColumnSchema).nonempty(),
  headerRowHeights: z.array(positiveNumber).optional(),
  bodyRowHeight: positiveNumber.optional(),
  headerStyle: styleValueSchema.optional(),
  bodyStyle: styleValueSchema.optional(),
  evenRowStyle: styleValueSchema.optional(),
  oddRowStyle: styleValueSchema.optional(),
  footerRows: z.array(tableFooterRowSchema).optional(),
  summaryStyle: styleValueSchema.optional(),
  rowHidden: functionSchema.optional(),
  border: z.unknown().optional(),
} as const;

// ─── Block & Workbook ─────────────────────────────────────────────────────────

const blockSchema = z.discriminatedUnion('type', [
  z.object({
    type: z.literal('grid'),
    rows: z.array(gridRowSchema),
  }),
  z
    .object({
      type: z.literal('table'),
      ...baseTableShape,
      data: z.array(recordSchema).optional(),
      groups: z.array(tableGroupSchema).optional(),
    })
    .refine((val) => (val.data !== undefined) !== (val.groups !== undefined), {
      message: "TableBlock must have exactly one of 'data' or 'groups'.",
    }),
]);

const workbookSchema = z.object({
  metadata: z
    .object({
      title: z.string().optional(),
      author: z.string().optional(),
      company: z.string().optional(),
      subject: z.string().optional(),
      keywords: z.array(z.string()).optional(),
    })
    .optional(),
  defaultStyle: styleObjectSchema.optional(),
  styles: z.record(z.string(), styleObjectSchema).optional(),
  sheets: z
    .array(
      z.object({
        id: nonEmptyString,
        name: nonEmptyString,
        freezePane: z
          .object({ rows: positiveInteger.optional(), columns: positiveInteger.optional() })
          .refine((p) => p.rows !== undefined || p.columns !== undefined, {
            message: 'must include rows or columns',
          })
          .optional(),
        blocks: z.array(blockSchema),
      }),
    )
    .nonempty('Workbook definition must include at least one sheet.'),
});

// ─── Public entry point ───────────────────────────────────────────────────────

export function validateWorkbookDefinition(workbook: WorkbookDefinition): void {
  const parsed = workbookSchema.safeParse(workbook);

  if (!parsed.success) {
    throw buildValidationError(parsed.error.issues);
  }

  validateSemantics(parsed.data as WorkbookDefinition);
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
      rejectDeprecatedFields(block, blockLabel);
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

// ─── Deprecated field detection ───────────────────────────────────────────────

const DEPRECATED_FIELDS = ['key', 'columnKey', 'startKey', 'endKey'] as const;

function rejectDeprecatedFields(value: unknown, rootLabel: string, path: string[] = []): void {
  if (!isPlainObject(value)) return;

  const label = path.length === 0 ? rootLabel : path.join(' ');

  for (const field of DEPRECATED_FIELDS) {
    if (Object.prototype.hasOwnProperty.call(value, field)) {
      throw new ValidationError(`${label} uses deprecated "${field}"; use id-based fields.`);
    }
  }

  for (const [key, child] of Object.entries(value)) {
    if (Array.isArray(child)) {
      child.forEach((item, i) => rejectDeprecatedFields(item, rootLabel, [...path, key, String(i)]));
    } else if (isPlainObject(child)) {
      rejectDeprecatedFields(child, rootLabel, [...path, key]);
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

// ─── Error helpers ────────────────────────────────────────────────────────────

function buildValidationError(issues: readonly ZodIssue[]): ValidationError {
  const issue = issues[0];
  if (!issue) return new ValidationError('Workbook definition is invalid.');
  return new ValidationError(formatIssueMessage(issue));
}

function formatIssueMessage(issue: ZodIssue): string {
  if (issue.message === 'Workbook definition must include at least one sheet.') return issue.message;
  const path = issue.path.length ? issue.path.join('.') : 'Workbook definition';
  return `${path} ${issue.message}`;
}

function blockPath(
  sheetId: string,
  block: WorkbookDefinition['sheets'][number]['blocks'][number],
  index: number,
): string {
  return `sheet "${sheetId}" > ${block.type} block ${index + 1}`;
}
