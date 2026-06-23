import { ReportEngineError } from "../core/errors";
import type {
  CellContent,
  FormulaBinaryOperator,
  FormulaDefinition,
  FormulaRangeReference,
} from "../core/types";
import type { RenderCell } from "./render-plan";

export interface FormulaCompileContext {
  resolveCellKey(key: string): string;
  resolveRangeKeys(startKey: string, endKey: string): string;
}

export function compileCellContent(
  content: CellContent | undefined,
  context?: FormulaCompileContext,
): Pick<RenderCell, "value" | "formula"> {
  if (content === undefined) {
    return {};
  }

  if (!isFormulaDefinition(content)) {
    return { value: content };
  }

  return { formula: compileFormula(content, context) };
}

export function compileFormula(
  formula: FormulaDefinition,
  context?: FormulaCompileContext,
): string {
  switch (formula.type) {
    case "raw":
      return compileRawFormula(formula.expression);
    case "literal":
      return compileLiteralFormula(formula.value);
    case "ref":
      return requireCompileContext(context).resolveCellKey(formula.key);
    case "range":
      return requireCompileContext(context).resolveRangeKeys(formula.startKey, formula.endKey);
    case "sum":
      return `SUM(${compileSumArguments(formula.range, formula.values, context)})`;
    case "round":
      return `ROUND(${compileFormula(formula.value, context)},${formula.digits})`;
    case "if":
      return [
        "IF(",
        compileFormula(formula.condition, context),
        ",",
        compileFormula(formula.whenTrue, context),
        ",",
        compileFormula(formula.whenFalse, context),
        ")",
      ].join("");
    case "call":
      return `${compileFunctionName(formula.name)}(${formula.args.map((arg) => compileFormula(arg, context)).join(",")})`;
    case "binary":
      return `(${compileFormula(formula.left, context)}${compileBinaryOperator(formula.operator)}${compileFormula(formula.right, context)})`;
    default:
      return assertNever(formula);
  }
}

export function isFormulaDefinition(value: unknown): value is FormulaDefinition {
  if (!isRecord(value) || typeof value.type !== "string") {
    return false;
  }

  return [
    "raw",
    "literal",
    "sum",
    "round",
    "if",
    "call",
    "binary",
    "range",
    "ref",
  ].includes(value.type);
}

export function createFormulaCompileContext(
  keyMap: Map<string, CellAddress>,
): FormulaCompileContext {
  return {
    resolveCellKey(key: string): string {
      const address = keyMap.get(key);

      if (!address) {
        throw new ReportEngineError(`Formula references unknown cell key "${key}".`);
      }

      return formatCellAddress(address);
    },
    resolveRangeKeys(startKey: string, endKey: string): string {
      const start = keyMap.get(startKey);
      const end = keyMap.get(endKey);

      if (!start) {
        throw new ReportEngineError(`Formula references unknown range start key "${startKey}".`);
      }

      if (!end) {
        throw new ReportEngineError(`Formula references unknown range end key "${endKey}".`);
      }

      if (end.row < start.row || end.column < start.column) {
        throw new ReportEngineError("Formula range end key must resolve after start key.");
      }

      return `${formatCellAddress(start)}:${formatCellAddress(end)}`;
    },
  };
}

export interface CellAddress {
  row: number;
  column: number;
}

export function formatCellAddress(address: CellAddress): string {
  assertPositiveInteger(address.row, "formula row");
  assertPositiveInteger(address.column, "formula column");

  return `${columnNumberToName(address.column)}${address.row}`;
}

function compileSumArguments(
  range: FormulaRangeReference | undefined,
  values: FormulaDefinition[] | undefined,
  context: FormulaCompileContext | undefined,
): string {
  const args = [
    ...(range ? [requireCompileContext(context).resolveRangeKeys(range.startKey, range.endKey)] : []),
    ...(values ? values.map((value) => compileFormula(value, context)) : []),
  ];

  if (args.length === 0) {
    throw new ReportEngineError("SUM formula must include a range or values.");
  }

  return args.join(",");
}

function compileRawFormula(expression: string): string {
  if (typeof expression !== "string" || expression.trim() === "") {
    throw new ReportEngineError("Raw formula expression must be a non-empty string.");
  }

  if (expression.trimStart().startsWith("=")) {
    throw new ReportEngineError("Formula expression must not start with '='.");
  }

  return expression;
}

function compileLiteralFormula(value: string | number | boolean | null): string {
  if (typeof value === "number") {
    return String(value);
  }

  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  if (value === null) {
    return "\"\"";
  }

  return `"${value.replace(/"/g, "\"\"")}"`;
}

function compileFunctionName(name: string): string {
  if (!/^[A-Za-z][A-Za-z0-9_.]*$/.test(name)) {
    throw new ReportEngineError("Formula function name must be a valid Excel function name.");
  }

  return name.toUpperCase();
}

function compileBinaryOperator(operator: FormulaBinaryOperator): string {
  const supportedOperators = new Set(["+", "-", "*", "/", ">", ">=", "<", "<=", "=", "<>"]);

  if (!supportedOperators.has(operator)) {
    throw new ReportEngineError(`Formula binary operator "${operator}" is not supported.`);
  }

  return operator;
}

function requireCompileContext(context: FormulaCompileContext | undefined): FormulaCompileContext {
  if (!context) {
    throw new ReportEngineError("Formula key references require a compile context.");
  }

  return context;
}

function columnNumberToName(column: number): string {
  let remaining = column;
  let name = "";

  while (remaining > 0) {
    remaining -= 1;
    name = String.fromCharCode(65 + (remaining % 26)) + name;
    remaining = Math.floor(remaining / 26);
  }

  return name;
}

function assertPositiveInteger(value: unknown, label: string): asserts value is number {
  if (typeof value !== "number" || !Number.isInteger(value) || value < 1) {
    throw new ReportEngineError(`${label} must be a positive integer.`);
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function assertNever(value: never): never {
  throw new ReportEngineError(`Unsupported formula type "${(value as FormulaDefinition).type}".`);
}
