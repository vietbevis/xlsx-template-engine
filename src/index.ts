export { defineWorkbook, isWorkbookDefinition } from './core/workbook';

export { f } from './formula';

export {
  CompileError,
  FormulaError,
  RenderError,
  ReportEngineError,
  ValidationError,
} from './core/errors';

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
  NamedRangeDefinition,
  NamedRangeFormulaDefinition,
  RawFormulaDefinition,
  RangeFormulaDefinition,
  RefFormulaDefinition,
  RepeatBlock,
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
  TableSectionCell,
  TableSectionCellAccessor,
  TableSectionCellContext,
  TableSectionRow,
  TableDataItem,
  TableTitleRow,
  TextBlock,
  TitleBlock,
  TypedFormulaDefinition,
  VlookupFormulaDefinition,
  TypedWorkbookDefinition,
  WorkbookDefinition,
  WorkbookMetadata,
} from './core/types';

export { compileWorkbookToRenderPlan, isFormulaDefinition } from './compiler';

export type { CompileWorkbookOptions, RenderContext, VariableScope } from './compiler';

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
  ResolvedNamedRange,
} from './compiler/render-plan';

export { renderWorkbook } from './renderer';

export type { WorkbookOutputAdapter, WorkbookRenderer, WorkbookRenderOptions } from './renderer';
