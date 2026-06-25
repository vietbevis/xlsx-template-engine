import type {
  FormulaExpression,
  FormulaRangeScope,
  FormulaTemplateDefinition,
  RangeFormulaDefinition,
  RefFormulaDefinition,
} from './types';

export function f(strings: TemplateStringsArray, ...exprs: FormulaExpression[]): FormulaTemplateDefinition {
  return { type: 'formula_template', strings, exprs };
}

f.ref = (id: string, sheetId?: string): RefFormulaDefinition => {
  return { type: 'ref', id, sheetId };
};

f.range = (
  startId: string,
  endId: string,
  options: { sheetId?: string; scope?: FormulaRangeScope } = {},
): RangeFormulaDefinition => {
  return { type: 'range', startId, endId, ...options };
};
