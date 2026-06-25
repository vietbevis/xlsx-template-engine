export { f } from './formula';

export { CompileError, FormulaError, RenderError, ReportEngineError, ValidationError } from './errors';

export type {
  Block,
  BinaryFormulaDefinition,
  CallFormulaDefinition,
  CellContent,
  CellStyleDefinition,
  CellValue,
  AverageFormulaDefinition,
  ConcatenateFormulaDefinition,
  CountAFormulaDefinition,
  CountFormulaDefinition,
  DividerBlock,
  FormulaBinaryOperator,
  FormulaDefinition,
  FormulaRangeReference,
  FormulaRangeScope,
  GridBlock,
  GridCell,
  GridRow,
  IfFormulaDefinition,
  IfErrorFormulaDefinition,
  LiteralFormulaDefinition,
  MaxFormulaDefinition,
  MinFormulaDefinition,
  RawFormulaDefinition,
  RangeFormulaDefinition,
  RefFormulaDefinition,
  RoundFormulaDefinition,
  SheetDefinition,
  SheetFreezePane,
  SpacerBlock,
  StyleReference,
  StyleRegistry,
  StyleValue,
  SumFormulaDefinition,
  TableBlock,
  TableBorderDefinition,
  TableColumn,
  TableFooterRow,
  TableGroup,
  TableGroupsBlock,
  TableSectionCell,
  TableSectionCellAccessor,
  TableSectionCellContext,
  TableSectionRow,
  TextBlock,
  TitleBlock,
  TypedFormulaDefinition,
  WorkbookDefinition,
  WorkbookMetadata,
} from './types';

export { compileWorkbookToRenderPlan } from './compile';
export { isFormulaDefinition } from './formula-engine';

export type { CompileWorkbookOptions } from './compile';
export type { RenderContext, VariableScope } from './variable-engine';

export type {
  RenderCell,
  RenderColumnVisibility,
  RenderColumnWidth,
  RenderCommand,
  RenderLink,
  RenderMergeRange,
  RenderPlan,
  RenderPlanSheet,
  RenderRow,
  RenderRowHeight,
  RenderRowVisibility,
  RenderSheetView,
} from './render-plan';

export { renderWorkbook } from './renderer';

export type { WorkbookOutputAdapter, WorkbookRenderer, WorkbookRenderOptions } from './renderer';
