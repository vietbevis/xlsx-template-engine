export { f } from './formula/formula-builder';
export { FormulaCompiler } from './formula/formula-compiler';

export { CompileError, FormulaError, RenderError, ReportEngineError, ValidationError } from './errors';

export type {
  Block,
  CellContent,
  CellStyleDefinition,
  CellValue,
  FormulaDefinition,
  FormulaExpression,
  FormulaRangeScope,
  FormulaTemplateDefinition,
  GridBlock,
  GridCell,
  GridRow,
  RangeFormulaDefinition,
  RefFormulaDefinition,
  SheetDefinition,
  StyleReference,
  StyleRegistry,
  StyleValue,
  TableBlock,
  TableBorderDefinition,
  TableColumn,
  TableFooterRow,
  TableGroup,
  TableSectionCell,
  TableSectionCellAccessor,
  TableSectionCellContext,
  TableSectionRow,
  WorkbookDefinition,
  WorkbookMetadata,
} from './types';

export type { CompileWorkbookOptions, WorkbookRenderer, WorkbookRenderOptions } from './compiler/workbook-compiler';
export { WorkbookCompiler, compileWorkbook, renderWorkbook } from './compiler/workbook-compiler';
export { dividerBlock, spacerBlock, textBlock } from './factories';
