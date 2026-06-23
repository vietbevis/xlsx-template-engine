export {
  defineWorkbook,
  isWorkbookDefinition,
} from "./core/workbook";

export {
  ReportEngineError,
} from "./core/errors";

export {
  createSheetRegistry,
} from "./core/sheet-registry";

export type {
  SheetRegistry,
} from "./core/sheet-registry";

export type {
  Block,
  BorderLineStyle,
  BorderSideStyleDefinition,
  BorderStyleDefinition,
  BinaryFormulaDefinition,
  CallFormulaDefinition,
  CellContent,
  CellStyleDefinition,
  CellValue,
  ColorStyleDefinition,
  FillPatternStyle,
  FillStyleDefinition,
  FormulaBinaryOperator,
  FontStyleDefinition,
  FormulaDefinition,
  FormulaRangeReference,
  GridBlock,
  GridCell,
  GridRow,
  HorizontalAlignmentStyle,
  IfFormulaDefinition,
  LiteralFormulaDefinition,
  RawFormulaDefinition,
  RangeFormulaDefinition,
  RefFormulaDefinition,
  RoundFormulaDefinition,
  SheetDefinition,
  SpacerBlock,
  StyleReference,
  StyleRegistry,
  SumFormulaDefinition,
  TableBlock,
  TableColumn,
  TextBlock,
  TitleBlock,
  VerticalAlignmentStyle,
  WorkbookDefinition,
  WorkbookMetadata,
} from "./core/types";

export {
  compileBlock,
  compileCellContent,
  compileFormula,
  compileWorkbookToRenderPlan,
  collectFormulaDependencies,
  createFormulaCompileContext,
  createFormulaKey,
  defaultBlockCompilerRegistry,
  formatCellAddress,
  formatCellReference,
  assertMergeDoesNotOverlap,
  interpolateCellValue,
  interpolateVariables,
  isFormulaDefinition,
  LayoutCursor,
  normalizeMergeRange,
  resolvePath,
  RenderPlanBuilder,
} from "./compiler";

export type {
  BlockCompiler,
  BlockCompilerRegistry,
  CellAddress,
  CompileWorkbookOptions,
  FormulaCompileContextOptions,
  FormulaDependencyGraph,
  FormulaCompileContext,
  MergeRange,
  RenderContext,
  SheetContext,
  VariableScope,
} from "./compiler";

export type {
  RenderCell,
  RenderColumnWidth,
  RenderCommand,
  RenderLink,
  RenderMergeRange,
  RenderPlan,
  RenderPlanSheet,
  RenderRow,
  RenderRowHeight,
} from "./compiler/render-plan";

export {
  renderWorkbook,
} from "./renderer";

export type {
  WorkbookOutputAdapter,
  WorkbookRenderer,
  WorkbookRenderOptions,
} from "./renderer";
