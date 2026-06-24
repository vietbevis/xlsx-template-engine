export { createSheetRegistry } from './core/sheet-registry';

export type { SheetRegistry } from './core/sheet-registry';

export {
  assertMergeDoesNotOverlap,
  collectFormulaDependencies,
  compileBlock,
  compileCellContent,
  compileFormula,
  createFormulaCompileContext,
  createFormulaId,
  defaultBlockCompilerRegistry,
  formatCellAddress,
  formatCellReference,
  interpolateCellValue,
  interpolateVariables,
  LayoutCursor,
  normalizeMergeRange,
  RenderPlanBuilder,
  resolvePath,
} from './compiler';

export type {
  BlockCompiler,
  BlockCompilerRegistry,
  CellAddress,
  FormulaCompileContext,
  FormulaCompileContextOptions,
  FormulaDependencyGraph,
  MergeRange,
  NormalizedMergeRange,
  RenderContext,
  SheetContext,
  VariableScope,
} from './compiler';
