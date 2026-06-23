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
  CellStyleDefinition,
  CellValue,
  ColorStyleDefinition,
  FillPatternStyle,
  FillStyleDefinition,
  FontStyleDefinition,
  GridBlock,
  GridCell,
  GridRow,
  HorizontalAlignmentStyle,
  SheetDefinition,
  SpacerBlock,
  StyleReference,
  StyleRegistry,
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
  compileWorkbookToRenderPlan,
  defaultBlockCompilerRegistry,
  assertMergeDoesNotOverlap,
  LayoutCursor,
  normalizeMergeRange,
  RenderPlanBuilder,
} from "./compiler";

export type {
  BlockCompiler,
  BlockCompilerRegistry,
  MergeRange,
  SheetContext,
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
