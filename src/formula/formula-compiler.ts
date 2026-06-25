import { FormulaError } from '../errors';
import { isPlainObject } from '../utils/common';
import { assertPositiveInteger } from '../utils/common';
import type { CellContent, FormulaDefinition, FormulaExpression, FormulaRangeScope, WriterCell } from '../types';

// ─── Public interfaces ────────────────────────────────────────────────────────

export interface FormulaCompileContext {
  resolveCellId(id: string, sheetId?: string): string;
  resolveRangeIds(startId: string, endId: string, sheetId?: string, scope?: FormulaRangeScope): string;
}

export interface CellAddress {
  row: number;
  column: number;
  sheetId?: string;
  sheetName?: string;
}

// ─── FormulaCompiler ──────────────────────────────────────────────────────────

export class FormulaCompiler {
  static compileCellContent(
    content: CellContent | undefined,
    context?: FormulaCompileContext,
  ): Pick<WriterCell, 'value' | 'formula'> {
    if (content === undefined) {
      return {};
    }

    if (!FormulaCompiler.isFormulaDefinition(content)) {
      return { value: content };
    }

    return { formula: FormulaCompiler.compileFormula(content, context) };
  }

  static isFormulaDefinition(value: unknown): value is FormulaDefinition {
    if (!isPlainObject(value) || typeof value.type !== 'string') {
      return false;
    }

    return ['formula_template', 'range', 'ref'].includes(value.type);
  }

  static formatCellAddress(address: CellAddress): string {
    assertPositiveInteger(address.row, 'formula row');
    assertPositiveInteger(address.column, 'formula column');

    return `${FormulaCompiler.columnNumberToName(address.column)}${address.row}`;
  }

  static formatCellReference(address: CellAddress, currentSheetId?: string): string {
    const localAddress = FormulaCompiler.formatCellAddress(address);

    if (!address.sheetId || address.sheetId === currentSheetId) {
      return localAddress;
    }

    if (!address.sheetName) {
      throw new FormulaError(`Formula reference for sheet "${address.sheetId}" is missing sheet name.`);
    }

    return `${FormulaCompiler.quoteSheetName(address.sheetName)}!${localAddress}`;
  }

  // ─── Private helpers ──────────────────────────────────────────────────────

  private static compileFormula(formula: FormulaDefinition, context?: FormulaCompileContext): string {
    switch (formula.type) {
      case 'formula_template':
        return formula.strings.reduce((result, str, i) => {
          const rawExpr = formula.exprs[i];
          const expr = rawExpr !== undefined ? FormulaCompiler.compileExpression(rawExpr, context) : '';
          return result + str + expr;
        }, '');
      case 'ref':
        return FormulaCompiler.requireCompileContext(context).resolveCellId(formula.id, formula.sheetId);
      case 'range':
        return FormulaCompiler.requireCompileContext(context).resolveRangeIds(
          formula.startId,
          formula.endId,
          formula.sheetId,
          formula.scope,
        );
      default:
        return FormulaCompiler.assertNever(formula);
    }
  }

  private static compileExpression(expr: FormulaExpression, context?: FormulaCompileContext): string {
    if (expr === null || expr === undefined) return '';
    if (typeof expr === 'string' || typeof expr === 'number' || typeof expr === 'boolean') return String(expr);
    if (FormulaCompiler.isFormulaDefinition(expr)) return FormulaCompiler.compileFormula(expr, context);
    throw new FormulaError('Unsupported formula expression type.');
  }

  private static columnNumberToName(column: number): string {
    let remaining = column;
    let name = '';

    while (remaining > 0) {
      remaining -= 1;
      name = String.fromCharCode(65 + (remaining % 26)) + name;
      remaining = Math.floor(remaining / 26);
    }

    return name;
  }

  private static quoteSheetName(sheetName: string): string {
    return `'${sheetName.replace(/'/g, "''")}'`;
  }

  private static requireCompileContext(context: FormulaCompileContext | undefined): FormulaCompileContext {
    if (!context) {
      throw new FormulaError('Formula id references require a compile context.');
    }

    return context;
  }

  private static assertNever(value: never): never {
    throw new FormulaError(`Unsupported formula type "${(value as FormulaDefinition).type}".`);
  }
}
