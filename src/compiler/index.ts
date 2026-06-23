import { ReportEngineError } from "../core/errors";
import type {
  Block,
  FormulaDefinition,
  SheetDefinition,
  WorkbookDefinition,
} from "../core/types";
import { validateWorkbookDefinition } from "../core/validation";
import { createSheetRegistry, type SheetRegistry } from "../core/sheet-registry";
import type { RenderPlan } from "./render-plan";
import { RenderPlanBuilder } from "./render-plan-builder";
import { LayoutCursor } from "./layout-cursor";
import { compileBlock } from "./block-compiler";
import {
  createFormulaKey,
  type CellAddress,
} from "./formula-engine";
import type { RenderContext } from "./variable-engine";

export interface CompileWorkbookOptions {
  context?: RenderContext;
}

export function compileWorkbookToRenderPlan(
  workbook: WorkbookDefinition,
  options: CompileWorkbookOptions = {},
): RenderPlan {
  validateWorkbookDefinition(workbook);
  const sheetRegistry = createSheetRegistry(workbook.sheets);
  collectFormulaDependencies(workbook, sheetRegistry);
  const formulaKeys = collectWorkbookFormulaKeys(workbook);

  const workbookContext = {
    ...(workbook.context ?? {}),
    ...(options.context ?? {}),
  };
  const builder = new RenderPlanBuilder(workbook.metadata, workbook.styles);

  for (const sheet of workbook.sheets) {
    builder.addSheet(sheet.id, sheet.name);
  }

  for (const sheet of workbook.sheets) {
    const context = {
      workbook,
      sheet,
      styles: workbook.styles,
      variables: {
        workbook: workbookContext,
        sheet: sheet.context,
      },
      formulaKeys,
    };
    const cursor = new LayoutCursor();

    for (const block of sheet.blocks) {
      compileBlock(block, context, cursor, builder);
    }
  }

  return builder.build();
}

export type FormulaDependencyGraph = ReadonlyMap<string, readonly string[]>;

export function collectFormulaDependencies(
  workbook: WorkbookDefinition,
  sheetRegistry: SheetRegistry = createSheetRegistry(workbook.sheets),
): FormulaDependencyGraph {
  const dependencies = new Map<string, Set<string>>();

  for (const sheet of workbook.sheets) {
    const sheetDependencies = new Set<string>();

    for (const formula of collectSheetFormulas(sheet)) {
      collectFormulaSheetIds(formula, sheet.id, sheetRegistry, sheetDependencies);
    }

    dependencies.set(sheet.id, sheetDependencies);
  }

  return new Map(
    Array.from(dependencies.entries()).map(([sheetId, sheetDependencies]) => [
      sheetId,
      Array.from(sheetDependencies),
    ]),
  );
}

function collectWorkbookFormulaKeys(workbook: WorkbookDefinition): Map<string, CellAddress> {
  const formulaKeys = new Map<string, CellAddress>();

  for (const sheet of workbook.sheets) {
    const cursor = new LayoutCursor();

    for (const block of sheet.blocks) {
      switch (block.type) {
        case "title":
        case "text":
          cursor.advanceRows();
          break;
        case "spacer":
          cursor.advanceRows(block.rows ?? 1);
          break;
        case "grid":
          cursor.advanceRows(collectGridFormulaKeys(formulaKeys, sheet, block, cursor));
          break;
        case "table":
          if (!Array.isArray(block.data)) {
            throw new ReportEngineError("Table async iterable data is not supported until streaming renderer phase 15.");
          }

          cursor.advanceRows(block.data.length + calculateTableHeaderDepth(block.columns));
          break;
        default:
          assertNever(block);
      }
    }
  }

  return formulaKeys;
}

function collectGridFormulaKeys(
  formulaKeys: Map<string, CellAddress>,
  sheet: SheetDefinition,
  block: Extract<Block, { type: "grid" }>,
  cursor: LayoutCursor,
): number {
  const occupied = new Set<string>();
  let rowExtent = block.rows.length;

  for (const [rowOffset, gridRow] of block.rows.entries()) {
    let columnOffset = 0;

    for (const cell of gridRow.cells) {
      while (occupied.has(gridOccupancyKey(rowOffset, columnOffset))) {
        columnOffset += 1;
      }

      const rowSpan = cell.rowSpan ?? 1;
      const colSpan = cell.colSpan ?? 1;
      const row = cursor.row + rowOffset;
      const column = cursor.column + columnOffset;

      if (cell.key) {
        registerWorkbookFormulaKey(formulaKeys, sheet, cell.key, row, column);
      }

      markGridOccupied(occupied, rowOffset, columnOffset, rowSpan, colSpan);
      rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
      columnOffset += colSpan;
    }
  }

  return rowExtent;
}

function registerWorkbookFormulaKey(
  formulaKeys: Map<string, CellAddress>,
  sheet: SheetDefinition,
  key: string,
  row: number,
  column: number,
): void {
  const registryKey = createFormulaKey(sheet.id, key);

  if (formulaKeys.has(registryKey)) {
    throw new ReportEngineError(`Duplicate formula cell key "${key}" in sheet "${sheet.id}".`);
  }

  formulaKeys.set(registryKey, {
    row,
    column,
    sheetId: sheet.id,
    sheetName: sheet.name,
  });
}

function collectSheetFormulas(sheet: SheetDefinition): FormulaDefinition[] {
  const formulas: FormulaDefinition[] = [];

  for (const block of sheet.blocks) {
    if (block.type === "grid") {
      for (const row of block.rows) {
        for (const cell of row.cells) {
          if (isFormulaObject(cell.value)) {
            formulas.push(cell.value);
          }
        }
      }
    }

    if (block.type === "table" && Array.isArray(block.data)) {
      for (const column of flattenColumns(block.columns)) {
        for (const row of block.data) {
          if (!column.accessor) {
            continue;
          }

          const value = column.accessor(row);

          if (isFormulaObject(value)) {
            formulas.push(value);
          }
        }
      }
    }
  }

  return formulas;
}

function collectFormulaSheetIds(
  formula: FormulaDefinition,
  currentSheetId: string,
  sheetRegistry: SheetRegistry,
  dependencies: Set<string>,
): void {
  if ("sheetId" in formula && formula.sheetId !== undefined) {
    if (!sheetRegistry.has(formula.sheetId)) {
      throw new ReportEngineError(`Formula references unknown sheetId "${formula.sheetId}".`);
    }

    if (formula.sheetId !== currentSheetId) {
      dependencies.add(formula.sheetId);
    }
  }

  switch (formula.type) {
    case "raw":
    case "literal":
    case "ref":
    case "range":
      return;
    case "sum":
      formula.values?.forEach((value) => collectFormulaSheetIds(value, currentSheetId, sheetRegistry, dependencies));
      return;
    case "round":
      collectFormulaSheetIds(formula.value, currentSheetId, sheetRegistry, dependencies);
      return;
    case "if":
      collectFormulaSheetIds(formula.condition, currentSheetId, sheetRegistry, dependencies);
      collectFormulaSheetIds(formula.whenTrue, currentSheetId, sheetRegistry, dependencies);
      collectFormulaSheetIds(formula.whenFalse, currentSheetId, sheetRegistry, dependencies);
      return;
    case "call":
      formula.args.forEach((arg) => collectFormulaSheetIds(arg, currentSheetId, sheetRegistry, dependencies));
      return;
    case "binary":
      collectFormulaSheetIds(formula.left, currentSheetId, sheetRegistry, dependencies);
      collectFormulaSheetIds(formula.right, currentSheetId, sheetRegistry, dependencies);
      return;
    default:
      assertNever(formula);
  }
}

function isFormulaObject(value: unknown): value is FormulaDefinition {
  return typeof value === "object" && value !== null && "type" in value;
}

type TableColumnNode = Extract<Block, { type: "table" }>["columns"][number];
type TableLeafColumn = TableColumnNode & { children?: undefined };

function flattenColumns(columns: TableColumnNode[]): TableLeafColumn[] {
  return columns.flatMap((column) => {
    if (column.children && column.children.length > 0) {
      return flattenColumns(column.children);
    }

    return [column as TableLeafColumn];
  });
}

function calculateTableHeaderDepth(columns: TableColumnNode[]): number {
  return Math.max(...columns.map((column) => {
    if (column.children && column.children.length > 0) {
      return 1 + calculateTableHeaderDepth(column.children);
    }

    return 1;
  }));
}

function markGridOccupied(
  occupied: Set<string>,
  rowOffset: number,
  columnOffset: number,
  rowSpan: number,
  colSpan: number,
): void {
  for (let row = rowOffset; row < rowOffset + rowSpan; row += 1) {
    for (let column = columnOffset; column < columnOffset + colSpan; column += 1) {
      occupied.add(gridOccupancyKey(row, column));
    }
  }
}

function gridOccupancyKey(rowOffset: number, columnOffset: number): string {
  return `${rowOffset}:${columnOffset}`;
}

function assertNever(value: never): never {
  throw new ReportEngineError(`Unsupported block type "${(value as Block).type}".`);
}

export { LayoutCursor } from "./layout-cursor";
export { RenderPlanBuilder } from "./render-plan-builder";
export {
  compileBlock,
  defaultBlockCompilerRegistry,
} from "./block-compiler";
export {
  compileCellContent,
  compileFormula,
  createFormulaCompileContext,
  createFormulaKey,
  formatCellAddress,
  formatCellReference,
  isFormulaDefinition,
} from "./formula-engine";
export {
  assertMergeDoesNotOverlap,
  normalizeMergeRange,
} from "./merge-engine";
export {
  interpolateCellValue,
  interpolateVariables,
  resolvePath,
} from "./variable-engine";

export type {
  BlockCompiler,
  BlockCompilerRegistry,
  SheetContext,
} from "./block-compiler";
export type {
  MergeRange,
} from "./merge-engine";
export type {
  CellAddress,
  FormulaCompileContext,
  FormulaCompileContextOptions,
} from "./formula-engine";
export type {
  RenderContext,
  VariableScope,
} from "./variable-engine";
