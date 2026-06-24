import { isFormulaDefinition } from './compiler/formula-engine';
import type {
  BinaryFormulaDefinition,
  FormulaBinaryOperator,
  FormulaDefinition,
  FormulaRangeScope,
  LiteralFormulaDefinition,
} from './core/types';

type CoercibleValue = FormulaDefinition | string | number | boolean | null;

export const f = {
  ref(id: string) {
    return { type: 'ref', id } as const;
  },
  xref(sheetId: string, id: string) {
    return { type: 'ref', sheetId, id } as const;
  },
  range(
    startId: string,
    endId: string,
    options: { sheetId?: string; scope?: FormulaRangeScope } = {},
  ) {
    return { type: 'range', startId, endId, ...options } as const;
  },
  namedRange(name: string) {
    return { type: 'namedRange', name } as const;
  },
  val(value: string | number | boolean | null): LiteralFormulaDefinition {
    return { type: 'literal', value };
  },
  raw(expression: string) {
    return { type: 'raw', expression } as const;
  },
  sumRange(
    startId: string,
    endId: string,
    options: { sheetId?: string; scope?: FormulaRangeScope } = {},
  ) {
    return { type: 'sum', range: { startId, endId, ...options } } as const;
  },
  sum(...values: FormulaDefinition[]) {
    return { type: 'sum', values } as const;
  },
  add(left: FormulaDefinition, right: CoercibleValue) {
    return binary('+', left, right);
  },
  sub(left: FormulaDefinition, right: CoercibleValue) {
    return binary('-', left, right);
  },
  mul(left: FormulaDefinition, right: CoercibleValue) {
    return binary('*', left, right);
  },
  div(left: FormulaDefinition, right: CoercibleValue) {
    return binary('/', left, right);
  },
  round(value: FormulaDefinition, digits: number) {
    return { type: 'round', value, digits } as const;
  },
  gt(left: FormulaDefinition, right: CoercibleValue) {
    return binary('>', left, right);
  },
  gte(left: FormulaDefinition, right: CoercibleValue) {
    return binary('>=', left, right);
  },
  lt(left: FormulaDefinition, right: CoercibleValue) {
    return binary('<', left, right);
  },
  lte(left: FormulaDefinition, right: CoercibleValue) {
    return binary('<=', left, right);
  },
  eq(left: FormulaDefinition, right: CoercibleValue) {
    return binary('=', left, right);
  },
  neq(left: FormulaDefinition, right: CoercibleValue) {
    return binary('<>', left, right);
  },
  if(condition: FormulaDefinition, whenTrue: CoercibleValue, whenFalse: CoercibleValue) {
    return {
      type: 'if',
      condition,
      whenTrue: coerce(whenTrue),
      whenFalse: coerce(whenFalse),
    } as const;
  },
  max(...values: FormulaDefinition[]) {
    return { type: 'max', values } as const;
  },
  min(...values: FormulaDefinition[]) {
    return { type: 'min', values } as const;
  },
  average(startId: string, endId: string, options: { sheetId?: string } = {}) {
    return { type: 'average', range: { startId, endId, ...options } } as const;
  },
  count(startId: string, endId: string, options: { sheetId?: string } = {}) {
    return { type: 'count', range: { startId, endId, ...options } } as const;
  },
  counta(startId: string, endId: string, options: { sheetId?: string } = {}) {
    return { type: 'counta', range: { startId, endId, ...options } } as const;
  },
  concat(...values: Array<FormulaDefinition | string>) {
    return { type: 'concatenate', values: values.map((value) => coerce(value)) } as const;
  },
  iferror(value: FormulaDefinition, fallback: CoercibleValue) {
    return { type: 'iferror', value, fallback: coerce(fallback) } as const;
  },
  vlookup(
    lookup: FormulaDefinition,
    rangeName: string,
    colIndex: number,
    exactMatch: boolean = true,
  ) {
    return { type: 'vlookup', lookup, rangeName, colIndex, exactMatch } as const;
  },
  call(name: string, ...args: FormulaDefinition[]) {
    return { type: 'call', name, args } as const;
  },
} as const;

function binary(
  operator: FormulaBinaryOperator,
  left: FormulaDefinition,
  right: CoercibleValue,
): BinaryFormulaDefinition {
  return {
    type: 'binary',
    operator,
    left,
    right: coerce(right),
  };
}

function coerce(value: CoercibleValue): FormulaDefinition {
  if (isFormulaDefinition(value)) {
    return value;
  }

  return { type: 'literal', value };
}
