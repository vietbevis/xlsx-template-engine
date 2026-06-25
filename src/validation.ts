import { ValidationError as ReportEngineError } from './errors';
import type { WorkbookDefinition } from './types';

export function validateWorkbookDefinition(workbook: WorkbookDefinition): void {
  if (!isPlainObject(workbook)) {
    throw new ReportEngineError('Workbook definition must be an object.');
  }

  if (!Array.isArray(workbook.sheets)) {
    throw new ReportEngineError('Workbook definition must include a sheets array.');
  }

  if (workbook.sheets.length === 0) {
    throw new ReportEngineError('Workbook definition must include at least one sheet.');
  }

  validateStyleObject(workbook.defaultStyle, 'Workbook defaultStyle');
  validateStyleRegistry(workbook.styles);

  const sheetIds = new Set<string>();
  const sheetNames = new Set<string>();

  for (const [index, sheet] of workbook.sheets.entries()) {
    if (!isPlainObject(sheet)) {
      throw new ReportEngineError(`Sheet at index ${index} must be an object.`);
    }

    if (typeof sheet.id !== 'string' || sheet.id.trim() === '') {
      throw new ReportEngineError(`Sheet at index ${index} must include a non-empty id.`);
    }

    if (typeof sheet.name !== 'string' || sheet.name.trim() === '') {
      throw new ReportEngineError(`Sheet "${sheet.id}" must include a non-empty name.`);
    }

    validateSheetName(sheet.name, sheet.id);

    if (!Array.isArray(sheet.blocks)) {
      throw new ReportEngineError(`Sheet "${sheet.id}" must include a blocks array.`);
    }

    validateFreezePane(sheet.freezePane, `Sheet "${sheet.id}" freezePane`);

    for (const [blockIndex, block] of sheet.blocks.entries()) {
      validateBlock(block, sheet.id, blockIndex, workbook.styles);
    }

    if (sheetIds.has(sheet.id)) {
      throw new ReportEngineError(`Duplicate sheet id "${sheet.id}".`);
    }

    const normalizedSheetName = sheet.name.trim().toLowerCase();

    if (sheetNames.has(normalizedSheetName)) {
      throw new ReportEngineError(`Duplicate sheet name "${sheet.name}".`);
    }

    sheetIds.add(sheet.id);
    sheetNames.add(normalizedSheetName);
  }
}

function validateFreezePane(value: unknown, label: string): void {
  if (value === undefined) {
    return;
  }

  if (!isPlainObject(value)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  validatePositiveInteger(value.rows, `${label} rows`);
  validatePositiveInteger(value.columns, `${label} columns`);

  if (value.rows === undefined && value.columns === undefined) {
    throw new ReportEngineError(`${label} must include rows or columns.`);
  }
}

function validateBlock(
  block: unknown,
  sheetId: string,
  blockIndex: number,
  styles: WorkbookDefinition['styles'],
): void {
  const path = createBlockPath(sheetId, block, blockIndex);

  if (!isPlainObject(block)) {
    throw new ReportEngineError(`${path} must be an object.`);
  }

  const blockType = block.type;

  if (typeof blockType !== 'string' || blockType.trim() === '') {
    throw new ReportEngineError(`${path} must include a non-empty type.`);
  }

  switch (blockType) {
    case 'title':
      validateTextBlock(block, path, styles);
      return;
    case 'text':
      validateTextBlock(block, path, styles);
      return;
    case 'spacer':
      if (block.rows !== undefined) {
        const rows = block.rows;

        if (typeof rows !== 'number' || !Number.isInteger(rows) || rows < 1) {
          throw new ReportEngineError(`${path} rows must be a positive integer.`);
        }
      }
      return;
    case 'divider':
      validatePositiveInteger(block.rows, `${path} rows`);
      validateStyleReference(block.style, styles, path);
      return;
    case 'grid':
      validateGridBlock(block, sheetId, blockIndex, styles);
      return;
    case 'table':
      validateTableBlock(block, sheetId, blockIndex, styles);
      return;
    case 'table-groups':
      validateTableGroupsBlock(block, sheetId, blockIndex, styles);
      return;
    default:
      throw new ReportEngineError(`Unknown block type "${blockType}" in sheet "${sheetId}".`);
  }
}

function validateTableBlock(
  block: Record<string, unknown>,
  sheetId: string,
  blockIndex: number,
  styles: WorkbookDefinition['styles'],
): void {
  const label = createBlockPath(sheetId, block, blockIndex);

  validateBaseTableBlock(block, sheetId, blockIndex, styles);

  if (!Array.isArray(block.data)) {
    throw new ReportEngineError(`${label} data must be an array.`);
  }

  if (block.groups !== undefined) {
    throw new ReportEngineError(`${label} must not include groups. Use type "table-groups" for grouped tables.`);
  }

  validateTableDataRows(block.data, `${label} data`);
}

function validateTableGroupsBlock(
  block: Record<string, unknown>,
  sheetId: string,
  blockIndex: number,
  styles: WorkbookDefinition['styles'],
): void {
  const label = createBlockPath(sheetId, block, blockIndex);

  validateBaseTableBlock(block, sheetId, blockIndex, styles);

  if (block.data !== undefined) {
    throw new ReportEngineError(`${label} must not include data. Put data inside groups[].data.`);
  }

  if (!Array.isArray(block.groups) || block.groups.length === 0) {
    throw new ReportEngineError(`${label} groups must be a non-empty array.`);
  }

  for (const [groupIndex, group] of block.groups.entries()) {
    validateTableGroup(group, `${label} groups ${groupIndex}`, styles);
  }
}

function validateBaseTableBlock(
  block: Record<string, unknown>,
  sheetId: string,
  blockIndex: number,
  styles: WorkbookDefinition['styles'],
): void {
  const label = createBlockPath(sheetId, block, blockIndex);

  if (!Array.isArray(block.columns) || block.columns.length === 0) {
    throw new ReportEngineError(`${label} columns must be a non-empty array.`);
  }

  if (block.headerRowHeights !== undefined) {
    if (!Array.isArray(block.headerRowHeights)) {
      throw new ReportEngineError(`${label} headerRowHeights must be an array.`);
    }

    const headerDepth = calculateTableHeaderDepth(block.columns);

    if (block.headerRowHeights.length > headerDepth) {
      throw new ReportEngineError(`${label} headerRowHeights must not exceed header row count.`);
    }

    for (const [rowIndex, height] of block.headerRowHeights.entries()) {
      validatePositiveNumber(height, `${label} headerRowHeights ${rowIndex}`);
    }
  }

  validatePositiveNumber(block.bodyRowHeight, `${label} bodyRowHeight`);
  validateStyleReference(block.headerStyle, styles, `${label} headerStyle`);
  validateStyleReference(block.bodyStyle, styles, `${label} bodyStyle`);
  validateStyleReference(block.evenRowStyle, styles, `${label} evenRowStyle`);
  validateStyleReference(block.oddRowStyle, styles, `${label} oddRowStyle`);
  validateStyleReference(block.summaryStyle, styles, `${label} summaryStyle`);

  if (block.rowHidden !== undefined && typeof block.rowHidden !== 'function') {
    throw new ReportEngineError(`${label} rowHidden must be a function.`);
  }

  if (block.titleRows !== undefined) {
    throw new ReportEngineError(
      `${label} must not include titleRows. Use separate title/text blocks before the table.`,
    );
  }

  if (block.preHeaderGrid !== undefined) {
    throw new ReportEngineError(`${label} must not include preHeaderGrid. Use a separate grid block before the table.`);
  }

  if (block.footerRows !== undefined) {
    if (!Array.isArray(block.footerRows)) {
      throw new ReportEngineError(`${label} footerRows must be an array.`);
    }

    for (const [rowIndex, row] of block.footerRows.entries()) {
      validateTableFooterRow(row, `${label} footerRows ${rowIndex}`, styles);
    }
  }

  for (const [columnIndex, column] of block.columns.entries()) {
    validateTableColumn(column, sheetId, blockIndex, [columnIndex], styles);
  }
}

function validateTableGroup(group: unknown, label: string, styles: WorkbookDefinition['styles']): void {
  if (!isPlainObject(group)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  if (group.headerRows !== undefined) {
    if (!Array.isArray(group.headerRows)) {
      throw new ReportEngineError(`${label} headerRows must be an array.`);
    }

    for (const [rowIndex, row] of group.headerRows.entries()) {
      validateTableSectionRow(row, `${label} headerRows ${rowIndex}`, styles);
    }
  }

  if (!Array.isArray(group.data)) {
    throw new ReportEngineError(`${label} data must be an array.`);
  }

  validateTableDataRows(group.data, `${label} data`);

  if (group.footerRows !== undefined) {
    if (!Array.isArray(group.footerRows)) {
      throw new ReportEngineError(`${label} footerRows must be an array.`);
    }

    for (const [rowIndex, row] of group.footerRows.entries()) {
      validateTableSectionRow(row, `${label} footerRows ${rowIndex}`, styles);
    }
  }
}

function validateTableDataRows(rows: readonly unknown[], label: string): void {
  for (const [itemIndex, item] of rows.entries()) {
    if (isLegacyTableSectionRow(item)) {
      throw new ReportEngineError(`${label} ${itemIndex} is a section row. Move it to headerRows or footerRows.`);
    }
  }
}

function validateTableFooterRow(row: unknown, label: string, styles: WorkbookDefinition['styles']): void {
  if (!isPlainObject(row)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  validatePositiveNumber(row.height, `${label} height`);
  validateStyleReference(row.style, styles, label);

  if (!Array.isArray(row.cells) || row.cells.length === 0) {
    throw new ReportEngineError(`${label} cells must be a non-empty array.`);
  }

  for (const [cellIndex, cell] of row.cells.entries()) {
    validateTableSectionCell(cell, `${label} cell ${cellIndex}`, styles);
  }
}

function validateTableSectionRow(row: unknown, label: string, styles: WorkbookDefinition['styles']): void {
  if (!isPlainObject(row)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  validatePositiveNumber(row.height, `${label} height`);
  validateStyleReference(row.style, styles, label);

  if (row.resetRows !== undefined && typeof row.resetRows !== 'boolean') {
    throw new ReportEngineError(`${label} resetRows must be a boolean.`);
  }

  if (row.hidden !== undefined && typeof row.hidden !== 'boolean') {
    throw new ReportEngineError(`${label} hidden must be a boolean.`);
  }

  if (!Array.isArray(row.cells) || row.cells.length === 0) {
    throw new ReportEngineError(`${label} cells must be a non-empty array.`);
  }

  for (const [cellIndex, cell] of row.cells.entries()) {
    validateTableSectionCell(cell, `${label} cell ${cellIndex}`, styles);
  }
}

function isLegacyTableSectionRow(value: unknown): value is Record<string, unknown> {
  return isPlainObject(value) && value.type === 'section';
}

function validateTableSectionCell(cell: unknown, label: string, styles: WorkbookDefinition['styles']): void {
  if (!isPlainObject(cell)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }
  rejectDeprecatedKeyFields(cell, label, ['key', 'columnKey']);

  if (cell.column !== undefined) {
    validatePositiveInteger(cell.column, `${label} column`);
  }

  validateOptionalId(cell.id, `${label} id`);
  validateOptionalKey(cell.columnId, `${label} columnId`);

  if (cell.column !== undefined && cell.columnId !== undefined) {
    throw new ReportEngineError(`${label} must not include both column and columnId.`);
  }

  if (cell.value !== undefined && typeof cell.value !== 'function') {
    validateCellContent(cell.value, `${label} value`);
  }

  if (cell.styleResolver !== undefined && typeof cell.styleResolver !== 'function') {
    throw new ReportEngineError(`${label} styleResolver must be a function.`);
  }

  if (cell.colSpan !== undefined && cell.colSpan !== 'remaining') {
    validatePositiveInteger(cell.colSpan, `${label} colSpan`);
  }

  validateStyleReference(cell.style, styles, label);
}

function validateTableColumn(
  column: unknown,
  sheetId: string,
  blockIndex: number,
  columnPath: number[],
  styles: WorkbookDefinition['styles'],
): void {
  const label = `Table column ${columnPath.join('.')} in block ${blockIndex} of sheet "${sheetId}"`;

  if (!isPlainObject(column)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  if (typeof column.title !== 'string' || column.title.trim() === '') {
    throw new ReportEngineError(`${label} title must be a non-empty string.`);
  }
  rejectDeprecatedKeyFields(column, label, ['key']);

  if (column.id !== undefined && (typeof column.id !== 'string' || column.id.trim() === '')) {
    throw new ReportEngineError(`${label} id must be a non-empty string.`);
  }

  if (column.accessor !== undefined && typeof column.accessor !== 'function') {
    throw new ReportEngineError(`${label} accessor must be a function.`);
  }

  if (column.hidden !== undefined && typeof column.hidden !== 'boolean') {
    throw new ReportEngineError(`${label} hidden must be a boolean.`);
  }

  if (column.styleResolver !== undefined && typeof column.styleResolver !== 'function') {
    throw new ReportEngineError(`${label} styleResolver must be a function.`);
  }

  const hasChildren = column.children !== undefined;

  if (hasChildren) {
    if (!Array.isArray(column.children) || column.children.length === 0) {
      throw new ReportEngineError(`${label} children must be a non-empty array.`);
    }

    if (column.id !== undefined || column.accessor !== undefined) {
      throw new ReportEngineError(`${label} with children must not include id or accessor.`);
    }

    for (const [childIndex, childColumn] of column.children.entries()) {
      validateTableColumn(childColumn, sheetId, blockIndex, [...columnPath, childIndex], styles);
    }
  } else if (column.childrenRowOffset !== undefined) {
    throw new ReportEngineError(`${label} childrenRowOffset requires children.`);
  } else if (column.id === undefined && column.accessor === undefined) {
    throw new ReportEngineError(`${label} leaf column must include an id or accessor.`);
  }

  validatePositiveInteger(column.childrenRowOffset, `${label} childrenRowOffset`);
  validatePositiveNumber(column.width, `${label} width`);
  validateStyleReference(column.style, styles, label);
  validateStyleReference(column.headerStyle, styles, `${label} headerStyle`);
  validateStyleReference(column.bodyStyle, styles, `${label} bodyStyle`);

  if (
    column.summary !== undefined &&
    column.summary !== 'sum' &&
    column.summary !== 'count' &&
    column.summary !== 'average'
  ) {
    validateCellContent(column.summary, `${label} summary`);
  }
}

function calculateTableHeaderDepth(columns: unknown[]): number {
  return Math.max(
    ...columns.map((column) => {
      if (!isPlainObject(column) || !Array.isArray(column.children) || column.children.length === 0) {
        return 1;
      }

      const childrenRowOffset = typeof column.childrenRowOffset === 'number' ? column.childrenRowOffset : 1;

      return childrenRowOffset + calculateTableHeaderDepth(column.children);
    }),
  );
}

function validateGridBlock(
  block: Record<string, unknown>,
  sheetId: string,
  blockIndex: number,
  styles: WorkbookDefinition['styles'],
): void {
  if (!Array.isArray(block.rows)) {
    throw new ReportEngineError(`Block ${blockIndex} in sheet "${sheetId}" grid rows must be an array.`);
  }

  for (const [rowIndex, row] of block.rows.entries()) {
    if (!isPlainObject(row)) {
      throw new ReportEngineError(`${createGridRowPath(sheetId, blockIndex, rowIndex)} must be an object.`);
    }

    validatePositiveNumber(row.height, `${createGridRowPath(sheetId, blockIndex, rowIndex)} height`);

    if (!Array.isArray(row.cells)) {
      throw new ReportEngineError(`${createGridRowPath(sheetId, blockIndex, rowIndex)} cells must be an array.`);
    }

    for (const [cellIndex, cell] of row.cells.entries()) {
      validateGridCell(cell, sheetId, blockIndex, rowIndex, cellIndex, styles);
    }
  }
}

function validateGridCell(
  cell: unknown,
  sheetId: string,
  blockIndex: number,
  rowIndex: number,
  cellIndex: number,
  styles: WorkbookDefinition['styles'],
): void {
  const label = createGridCellPath(sheetId, blockIndex, rowIndex, cellIndex, cell);

  if (!isPlainObject(cell)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }
  rejectDeprecatedKeyFields(cell, label, ['key']);

  validateOptionalId(cell.id, `${label} id`);

  if (cell.value !== undefined) {
    validateCellContent(cell.value, `${label} value`);
  }

  validateColSpan(cell.colSpan, `${label} colSpan`);
  validatePositiveInteger(cell.rowSpan, `${label} rowSpan`);
  validatePositiveNumber(cell.width, `${label} width`);
  validateStyleReference(cell.style, styles, label);

  if (cell.styleResolver !== undefined && typeof cell.styleResolver !== 'function') {
    throw new ReportEngineError(`${label} styleResolver must be a function.`);
  }
}

function validateStyleRegistry(styles: unknown): void {
  if (styles === undefined) {
    return;
  }

  if (!isPlainObject(styles)) {
    throw new ReportEngineError('Workbook styles must be an object.');
  }

  for (const [styleName, style] of Object.entries(styles)) {
    if (styleName.trim() === '') {
      throw new ReportEngineError('Workbook style names must be non-empty.');
    }

    validateStyleObject(style, `Style "${styleName}"`);
  }
}

function validateStyleReference(styleValue: unknown, styles: WorkbookDefinition['styles'], label: string): void {
  if (styleValue === undefined) {
    return;
  }

  if (typeof styleValue !== 'string') {
    if (!isPlainObject(styleValue)) {
      throw new ReportEngineError(`${label} style must be a non-empty string or style object.`);
    }

    return;
  }

  if (styleValue.trim() === '') {
    throw new ReportEngineError(`${label} style must be a non-empty string or style object.`);
  }

  if (!styles || !Object.prototype.hasOwnProperty.call(styles, styleValue)) {
    throw new ReportEngineError(`${label} references unknown style "${styleValue}".`);
  }
}

function validateStyleObject(style: unknown, label: string): void {
  if (style === undefined) {
    return;
  }

  if (!isPlainObject(style)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }
}

function validateTextBlock(block: Record<string, unknown>, label: string, styles: WorkbookDefinition['styles']): void {
  if (typeof block.text !== 'string') {
    throw new ReportEngineError(`${label} text must be a string.`);
  }

  validatePositiveNumber(block.height, `${label} height`);
  validateColSpan(block.colSpan, `${label} colSpan`);
  validateStyleReference(block.style, styles, label);
}

function validateColSpan(value: unknown, label: string): void {
  if (value === 'remaining') {
    return;
  }

  validatePositiveInteger(value, label);
}

function validatePositiveInteger(value: unknown, label: string): void {
  if (value === undefined) {
    return;
  }

  if (typeof value !== 'number' || !Number.isInteger(value) || value < 1) {
    throw new ReportEngineError(`${label} must be a positive integer.`);
  }
}

function validatePositiveNumber(value: unknown, label: string): void {
  if (value === undefined) {
    return;
  }

  if (typeof value !== 'number' || value <= 0) {
    throw new ReportEngineError(`${label} must be greater than 0.`);
  }
}

function validateCellContent(value: unknown, label: string): void {
  if (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean' ||
    value instanceof Date
  ) {
    return;
  }

  if (isPlainObject(value) && typeof value.type === 'string') {
    validateFormulaDefinition(value, label);
    return;
  }

  if (isPlainObject(value)) {
    return;
  }

  throw new ReportEngineError(`${label} must be a valid ExcelJS cell value or formula definition.`);
}

function validateFormulaDefinition(value: Record<string, unknown>, label: string): void {
  switch (value.type) {
    case 'raw':
      if (typeof value.expression !== 'string' || value.expression.trim() === '') {
        throw new ReportEngineError(`${label} raw formula expression must be a non-empty string.`);
      }

      if (value.expression.trimStart().startsWith('=')) {
        throw new ReportEngineError(`${label} formula expression must not start with '='.`);
      }
      return;
    case 'literal':
      if (
        value.value !== null &&
        typeof value.value !== 'string' &&
        typeof value.value !== 'number' &&
        typeof value.value !== 'boolean'
      ) {
        throw new ReportEngineError(`${label} literal formula value must be a string, number, boolean, or null.`);
      }
      return;
    case 'sum':
      if (value.range === undefined && value.values === undefined) {
        throw new ReportEngineError(`${label} sum formula must include a range or values.`);
      }

      if (value.range !== undefined) {
        if (!isPlainObject(value.range)) {
          throw new ReportEngineError(`${label} sum formula range must be an object.`);
        }

        validateFormulaRangeReference(value.range, `${label} sum formula range`);
      }

      if (value.values !== undefined) {
        if (!Array.isArray(value.values) || value.values.length === 0) {
          throw new ReportEngineError(`${label} sum formula values must be a non-empty array.`);
        }

        for (const [index, child] of value.values.entries()) {
          validateNestedFormulaDefinition(child, `${label} sum formula value ${index}`);
        }
      }
      return;
    case 'round':
      validateNestedFormulaDefinition(value.value, `${label} round formula value`);

      if (typeof value.digits !== 'number' || !Number.isInteger(value.digits) || value.digits < 0) {
        throw new ReportEngineError(`${label} round formula digits must be a non-negative integer.`);
      }
      return;
    case 'if':
      validateNestedFormulaDefinition(value.condition, `${label} if formula condition`);
      validateNestedFormulaDefinition(value.whenTrue, `${label} if formula whenTrue`);
      validateNestedFormulaDefinition(value.whenFalse, `${label} if formula whenFalse`);
      return;
    case 'call':
      if (typeof value.name !== 'string' || !/^[A-Za-z][A-Za-z0-9_.]*$/.test(value.name)) {
        throw new ReportEngineError(`${label} call formula name must be a valid Excel function name.`);
      }

      if (!Array.isArray(value.args)) {
        throw new ReportEngineError(`${label} call formula args must be an array.`);
      }

      for (const [index, child] of value.args.entries()) {
        validateNestedFormulaDefinition(child, `${label} call formula arg ${index}`);
      }
      return;
    case 'max':
    case 'min':
      validateFormulaArray(value.values, `${label} ${value.type} formula values`);
      return;
    case 'average':
    case 'count':
    case 'counta':
      if (!isPlainObject(value.range)) {
        throw new ReportEngineError(`${label} ${value.type} formula range must be an object.`);
      }

      validateFormulaRangeReference(value.range, `${label} ${value.type} formula range`);
      return;
    case 'concatenate':
      validateFormulaArray(value.values, `${label} concatenate formula values`);
      return;
    case 'iferror':
      validateNestedFormulaDefinition(value.value, `${label} iferror formula value`);
      validateNestedFormulaDefinition(value.fallback, `${label} iferror formula fallback`);
      return;
    case 'binary':
      if (
        typeof value.operator !== 'string' ||
        !['+', '-', '*', '/', '>', '>=', '<', '<=', '=', '<>'].includes(value.operator)
      ) {
        throw new ReportEngineError(`${label} binary formula operator is not supported.`);
      }

      validateNestedFormulaDefinition(value.left, `${label} binary formula left`);
      validateNestedFormulaDefinition(value.right, `${label} binary formula right`);
      return;
    case 'range':
      validateFormulaRangeReference(value, `${label} range formula`);
      return;
    case 'ref':
      validateOptionalKey(value.sheetId, `${label} ref formula sheetId`);
      rejectDeprecatedKeyFields(value, `${label} ref formula`, ['key']);
      validateRequiredKey(value.id, `${label} ref formula id`);
      return;
    default:
      throw new ReportEngineError(`${label} formula type is not supported.`);
  }
}

function validateFormulaArray(value: unknown, label: string): void {
  if (!Array.isArray(value) || value.length === 0) {
    throw new ReportEngineError(`${label} must be a non-empty array.`);
  }

  for (const [index, child] of value.entries()) {
    validateNestedFormulaDefinition(child, `${label} ${index}`);
  }
}

function validateNestedFormulaDefinition(value: unknown, label: string): void {
  if (!isPlainObject(value) || typeof value.type !== 'string') {
    throw new ReportEngineError(`${label} must be a formula definition.`);
  }

  validateFormulaDefinition(value, label);
}

function validateFormulaRangeReference(value: Record<string, unknown>, label: string): void {
  rejectDeprecatedKeyFields(value, label, ['startKey', 'endKey']);
  validateOptionalKey(value.sheetId, `${label} sheetId`);
  validateRequiredKey(value.startId, `${label} startId`);
  validateRequiredKey(value.endId, `${label} endId`);

  if (value.scope !== undefined && value.scope !== 'currentRows' && value.scope !== 'allRows') {
    throw new ReportEngineError(`${label} scope must be currentRows or allRows.`);
  }

  if (value.sheetId !== undefined && value.scope !== undefined) {
    throw new ReportEngineError(`${label} must not include both sheetId and scope.`);
  }
}

function validateOptionalKey(value: unknown, label: string): void {
  if (value === undefined) {
    return;
  }

  validateRequiredKey(value, label);
}

function validateOptionalId(value: unknown, label: string): void {
  validateOptionalKey(value, label);
}

function validateRequiredKey(value: unknown, label: string): void {
  if (typeof value !== 'string' || value.trim() === '') {
    throw new ReportEngineError(`${label} must be a non-empty string.`);
  }
}

function rejectDeprecatedKeyFields(value: Record<string, unknown>, label: string, fields: readonly string[]): void {
  for (const field of fields) {
    if (Object.prototype.hasOwnProperty.call(value, field)) {
      throw new ReportEngineError(`${label} uses deprecated "${field}"; use id-based fields.`);
    }
  }
}

function createBlockPath(sheetId: string, block: unknown, blockIndex: number): string {
  const type = isPlainObject(block) && typeof block.type === 'string' ? block.type : 'block';
  return `sheet "${sheetId}" > ${type} block ${blockIndex + 1}`;
}

function createGridRowPath(sheetId: string, blockIndex: number, rowIndex: number): string {
  return `sheet "${sheetId}" > grid block ${blockIndex + 1} > row ${rowIndex + 1}`;
}

function createGridCellPath(
  sheetId: string,
  blockIndex: number,
  rowIndex: number,
  cellIndex: number,
  cell: unknown,
): string {
  const base = `${createGridRowPath(sheetId, blockIndex, rowIndex)} > cell`;

  if (isPlainObject(cell)) {
    if (typeof cell.id === 'string' && cell.id.trim() !== '') {
      return `${base} "${cell.id}"`;
    }

    if (typeof cell.value === 'string' && cell.value.trim() !== '') {
      return `${base} "${cell.value}"`;
    }
  }

  return `${base} ${cellIndex + 1}`;
}

function validateSheetName(sheetName: string, sheetId: string): void {
  if (sheetName.length > 31) {
    throw new ReportEngineError(`Sheet "${sheetId}" name must be 31 characters or fewer.`);
  }

  if (/[:\\/?*[\]]/.test(sheetName)) {
    throw new ReportEngineError(`Sheet "${sheetId}" name contains characters Excel does not allow.`);
  }
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
