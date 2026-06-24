export { defineWorkbook, isWorkbookDefinition } from './core/workbook';

export { ReportEngineError } from './core/errors';

export type {
  Block,
  BinaryFormulaDefinition,
  CallFormulaDefinition,
  CellContent,
  CellStyleDefinition,
  CellValue,
  FormulaBinaryOperator,
  FormulaDefinition,
  FormulaRangeReference,
  FormulaRangeScope,
  GridBlock,
  GridCell,
  GridRow,
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
  StyleValue,
  SumFormulaDefinition,
  TableBlock,
  TableBorderDefinition,
  TableColumn,
  TableSectionCell,
  TableSectionCellAccessor,
  TableSectionCellContext,
  TableSectionRow,
  TableDataItem,
  TableTitleRow,
  TextBlock,
  TitleBlock,
  TypedFormulaDefinition,
  TypedWorkbookDefinition,
  WorkbookDefinition,
  WorkbookMetadata,
} from './core/types';

export { compileWorkbookToRenderPlan, isFormulaDefinition } from './compiler';

export type { CompileWorkbookOptions, RenderContext, VariableScope } from './compiler';

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
} from './compiler/render-plan';

export { renderWorkbook } from './renderer';

export type { WorkbookOutputAdapter, WorkbookRenderer, WorkbookRenderOptions } from './renderer';
