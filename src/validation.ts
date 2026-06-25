import { z, type ZodIssue } from 'zod';
import { ValidationError } from './errors';
import { calculateTableHeaderDepth } from './helpers/table';
import type { WorkbookDefinition } from './types';

const nonEmptyString = z.string().trim().min(1);
const positiveNumber = z.number().positive();
const positiveInteger = z.number().int().positive();
const styleObjectSchema = z.record(z.string(), z.unknown());
const styleValueSchema = z.union([nonEmptyString, styleObjectSchema]);
const colSpanSchema = z.union([positiveInteger, z.literal('remaining')]);
const functionSchema = z.custom<(...args: never[]) => unknown>((value) => typeof value === 'function', {
  message: 'must be a function',
});

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
    z.object({ type: z.literal('ref'), sheetId: nonEmptyString.optional(), id: nonEmptyString }),
  ]),
);

const formulaRangeSchema = z
  .object({
    sheetId: nonEmptyString.optional(),
    startId: nonEmptyString,
    endId: nonEmptyString,
    scope: z.enum(['currentRows', 'allRows']).optional(),
  })
  .refine((value) => !(value.sheetId && value.scope), {
    message: 'must not include both sheetId and scope',
  });

const cellContentSchema = z.custom((value) => {
  if (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean' ||
    value instanceof Date
  ) {
    return true;
  }

  if (isRecord(value) && typeof value.type === 'string') {
    return formulaSchema.safeParse(value).success;
  }

  return isRecord(value);
});

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

const tableSectionRowSchema = z.object({
  resetRows: z.boolean().optional(),
  hidden: z.boolean().optional(),
  height: positiveNumber.optional(),
  style: styleValueSchema.optional(),
  cells: z.array(tableSectionCellSchema).nonempty(),
});

const tableFooterRowSchema = z.object({
  height: positiveNumber.optional(),
  style: styleValueSchema.optional(),
  cells: z.array(tableSectionCellSchema).nonempty(),
});

const tableColumnSchema: z.ZodType<unknown> = z.lazy(() =>
  z
    .object({
      title: nonEmptyString,
      id: nonEmptyString.optional(),
      accessor: functionSchema.optional(),
      children: z.array(tableColumnSchema).nonempty().optional(),
      childrenRowOffset: positiveInteger.optional(),
      width: positiveNumber.optional(),
      hidden: z.boolean().optional(),
      style: styleValueSchema.optional(),
      headerStyle: styleValueSchema.optional(),
      bodyStyle: styleValueSchema.optional(),
      styleResolver: functionSchema.optional(),
      summary: z.union([z.enum(['sum', 'count', 'average']), formulaSchema]).optional(),
    })
    .superRefine((column, ctx) => {
      if (column.children) {
        if (column.id !== undefined || column.accessor !== undefined) {
          ctx.addIssue({ code: 'custom', message: 'with children must not include id or accessor' });
        }
        return;
      }

      if (column.childrenRowOffset !== undefined) {
        ctx.addIssue({ code: 'custom', message: 'childrenRowOffset requires children' });
      }

      if (column.id === undefined && column.accessor === undefined) {
        ctx.addIssue({ code: 'custom', message: 'leaf column must include an id or accessor' });
      }
    }),
);

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
};

const blockSchema = z.discriminatedUnion('type', [
  z.object({
    type: z.literal('grid'),
    rows: z.array(gridRowSchema),
    context: z.record(z.string(), z.unknown()).optional(),
  }),
  z
    .object({
      type: z.literal('table'),
      ...baseTableShape,
      data: z.array(z.record(z.string(), z.unknown())).optional(),
      groups: z
        .array(
          z.object({
            headerRows: z.array(tableSectionRowSchema).optional(),
            data: z.array(z.record(z.string(), z.unknown())),
            footerRows: z.array(tableSectionRowSchema).optional(),
          }),
        )
        .optional(),
      context: z.record(z.string(), z.unknown()).optional(),
    })
    .refine(
      (val) => {
        const hasData = val.data !== undefined;
        const hasGroups = val.groups !== undefined;
        return (hasData && !hasGroups) || (!hasData && hasGroups);
      },
      { message: "TableBlock must have exactly one of 'data' or 'groups'." },
    ),
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
  context: z.record(z.string(), z.unknown()).optional(),
  sheets: z
    .array(
      z.object({
        id: nonEmptyString,
        name: nonEmptyString,
        context: z.record(z.string(), z.unknown()).optional(),
        freezePane: z
          .object({ rows: positiveInteger.optional(), columns: positiveInteger.optional() })
          .refine((value) => value.rows !== undefined || value.columns !== undefined, {
            message: 'must include rows or columns',
          })
          .optional(),
        blocks: z.array(blockSchema),
      }),
    )
    .nonempty('Workbook definition must include at least one sheet.'),
});

export function validateWorkbookDefinition(workbook: WorkbookDefinition): void {
  const parsed = workbookSchema.safeParse(workbook);

  if (!parsed.success) {
    throw createValidationError(parsed.error.issues);
  }

  validateWorkbookSemantics(parsed.data as WorkbookDefinition);
}

function validateWorkbookSemantics(workbook: WorkbookDefinition): void {
  validateStyleRegistryNames(workbook.styles);

  const sheetIds = new Set<string>();
  const sheetNames = new Set<string>();

  for (const [sheetIndex, sheet] of workbook.sheets.entries()) {
    validateSheetName(sheet.name, sheet.id);

    if (sheetIds.has(sheet.id)) {
      throw new ValidationError(`Duplicate sheet id "${sheet.id}".`);
    }

    const normalizedSheetName = sheet.name.trim().toLowerCase();

    if (sheetNames.has(normalizedSheetName)) {
      throw new ValidationError(`Duplicate sheet name "${sheet.name}".`);
    }

    sheetIds.add(sheet.id);
    sheetNames.add(normalizedSheetName);

    for (const [blockIndex, block] of sheet.blocks.entries()) {
      validateBlockSemantics(block, sheet.id, blockIndex, workbook.styles);
      rejectDeprecatedBlockFields(block, sheet.id, blockIndex);
    }
  }
}

function validateBlockSemantics(
  block: WorkbookDefinition['sheets'][number]['blocks'][number],
  sheetId: string,
  blockIndex: number,
  styles: WorkbookDefinition['styles'],
): void {
  validateStyleReferences(block, createBlockPath(sheetId, block, blockIndex), styles);

  if (block.type === 'table') {
    const label = createBlockPath(sheetId, block, blockIndex);

    if (block.headerRowHeights && block.headerRowHeights.length > calculateTableHeaderDepth(block.columns)) {
      throw new ValidationError(`${label} headerRowHeights must not exceed header row count.`);
    }

    validateTableDataRows(block.data ? block.data : block.groups!.flatMap((group) => group.data), `${label} data`);
  }
}

function validateStyleReferences(value: unknown, label: string, styles: WorkbookDefinition['styles']): void {
  if (!isRecord(value)) {
    return;
  }

  for (const key of ['style', 'headerStyle', 'bodyStyle', 'evenRowStyle', 'oddRowStyle', 'summaryStyle'] as const) {
    const style = value[key];

    if (typeof style === 'string' && (!styles || !Object.prototype.hasOwnProperty.call(styles, style))) {
      throw new ValidationError(`${label} ${key} references unknown style "${style}".`);
    }
  }

  for (const [key, child] of Object.entries(value)) {
    if (key === 'styles') {
      continue;
    }

    if (Array.isArray(child)) {
      child.forEach((item, index) => validateStyleReferences(item, `${label} ${key} ${index}`, styles));
    } else if (isRecord(child) && typeof child.type !== 'string') {
      validateStyleReferences(child, `${label} ${key}`, styles);
    }
  }
}

function validateStyleRegistryNames(styles: WorkbookDefinition['styles']): void {
  for (const styleName of Object.keys(styles ?? {})) {
    if (styleName.trim() === '') {
      throw new ValidationError('Workbook style names must be non-empty.');
    }
  }
}

function validateTableDataRows(rows: readonly unknown[], label: string): void {
  for (const [itemIndex, item] of rows.entries()) {
    if (isRecord(item) && item.type === 'section') {
      throw new ValidationError(`${label} ${itemIndex} is a section row. Move it to headerRows or footerRows.`);
    }
  }
}

function rejectDeprecatedBlockFields(
  block: WorkbookDefinition['sheets'][number]['blocks'][number],
  sheetId: string,
  blockIndex: number,
): void {
  visitRecords(block, (record, path) => {
    const label = path.length === 0 ? createBlockPath(sheetId, block, blockIndex) : path.join(' ');

    rejectDeprecatedFields(record, label, ['key', 'columnKey', 'startKey', 'endKey']);
  });
}

function rejectDeprecatedFields(value: Record<string, unknown>, label: string, fields: readonly string[]): void {
  for (const field of fields) {
    if (Object.prototype.hasOwnProperty.call(value, field)) {
      throw new ValidationError(`${label} uses deprecated "${field}"; use id-based fields.`);
    }
  }
}

function visitRecords(
  value: unknown,
  visitor: (record: Record<string, unknown>, path: string[]) => void,
  path: string[] = [],
): void {
  if (!isRecord(value)) {
    return;
  }

  visitor(value, path);

  for (const [key, child] of Object.entries(value)) {
    if (Array.isArray(child)) {
      child.forEach((item, index) => visitRecords(item, visitor, [...path, key, String(index)]));
    } else if (isRecord(child)) {
      visitRecords(child, visitor, [...path, key]);
    }
  }
}

function validateSheetName(sheetName: string, sheetId: string): void {
  if (sheetName.length > 31) {
    throw new ValidationError(`Sheet "${sheetId}" name must be 31 characters or fewer.`);
  }

  if (/[:\\/?*[\]]/.test(sheetName)) {
    throw new ValidationError(`Sheet "${sheetId}" name contains characters Excel does not allow.`);
  }
}

function createValidationError(issues: readonly ZodIssue[]): ValidationError {
  const issue = issues[0];

  if (!issue) {
    return new ValidationError('Workbook definition is invalid.');
  }

  return new ValidationError(formatIssueMessage(issue));
}

function formatIssueMessage(issue: ZodIssue): string {
  if (issue.message === 'Workbook definition must include at least one sheet.') {
    return issue.message;
  }

  const path = issue.path.length ? issue.path.join('.') : 'Workbook definition';
  return `${path} ${issue.message}`;
}

function createBlockPath(
  sheetId: string,
  block: WorkbookDefinition['sheets'][number]['blocks'][number],
  blockIndex: number,
): string {
  return `sheet "${sheetId}" > ${block.type} block ${blockIndex + 1}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
