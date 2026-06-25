import { FormulaError } from './errors';
import { assertPositiveInteger, isPlainObject } from './helpers/utils';
import type { CellContent, FormulaDefinition, FormulaExpression, FormulaRangeScope, WriterCell } from './types';

export interface FormulaCompileContext {
  resolveCellId(id: string, sheetId?: string): string;
  resolveRangeIds(startId: string, endId: string, sheetId?: string, scope?: FormulaRangeScope): string;
}

export function compileCellContent(
  content: CellContent | undefined,
  context?: FormulaCompileContext,
): Pick<WriterCell, 'value' | 'formula'> {
  if (content === undefined) {
    return {};
  }

  if (!isFormulaDefinition(content)) {
    return { value: content };
  }

  return { formula: compileFormula(content, context) };
}

export function compileFormula(formula: FormulaDefinition, context?: FormulaCompileContext): string {
  switch (formula.type) {
    case 'formula_template':
      return formula.strings.reduce((result, str, i) => {
        const rawExpr = formula.exprs[i];
        const expr = rawExpr !== undefined ? compileExpression(rawExpr, context) : '';
        return result + str + expr;
      }, '');
    case 'ref':
      return requireCompileContext(context).resolveCellId(formula.id, formula.sheetId);
    case 'range':
      return requireCompileContext(context).resolveRangeIds(
        formula.startId,
        formula.endId,
        formula.sheetId,
        formula.scope,
      );
    default:
      return assertNever(formula);
  }
}

function compileExpression(expr: FormulaExpression, context?: FormulaCompileContext): string {
  if (expr === null || expr === undefined) return '';
  if (typeof expr === 'string' || typeof expr === 'number' || typeof expr === 'boolean') return String(expr);
  if (isFormulaDefinition(expr)) return compileFormula(expr, context);
  throw new FormulaError('Unsupported formula expression type.');
}

export function isFormulaDefinition(value: unknown): value is FormulaDefinition {
  if (!isPlainObject(value) || typeof value.type !== 'string') {
    return false;
  }

  return ['formula_template', 'range', 'ref'].includes(value.type);
}

export interface CellAddress {
  row: number;
  column: number;
  sheetId?: string;
  sheetName?: string;
}

export function formatCellAddress(address: CellAddress): string {
  assertPositiveInteger(address.row, 'formula row');
  assertPositiveInteger(address.column, 'formula column');

  return `${columnNumberToName(address.column)}${address.row}`;
}

export function formatCellReference(address: CellAddress, currentSheetId?: string): string {
  const localAddress = formatCellAddress(address);

  if (!address.sheetId || address.sheetId === currentSheetId) {
    return localAddress;
  }

  if (!address.sheetName) {
    throw new FormulaError(`Formula reference for sheet "${address.sheetId}" is missing sheet name.`);
  }

  return `${quoteSheetName(address.sheetName)}!${localAddress}`;
}

function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replace(/'/g, "''")}'`;
}

function requireCompileContext(context: FormulaCompileContext | undefined): FormulaCompileContext {
  if (!context) {
    throw new FormulaError('Formula id references require a compile context.');
  }

  return context;
}

function columnNumberToName(column: number): string {
  let remaining = column;
  let name = '';

  while (remaining > 0) {
    remaining -= 1;
    name = String.fromCharCode(65 + (remaining % 26)) + name;
    remaining = Math.floor(remaining / 26);
  }

  return name;
}

function assertNever(value: never): never {
  throw new FormulaError(`Unsupported formula type "${(value as FormulaDefinition).type}".`);
}
