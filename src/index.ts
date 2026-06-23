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
  CellValue,
  SheetDefinition,
  StyleRegistry,
  WorkbookDefinition,
  WorkbookMetadata,
} from "./core/types";

export {
  compileWorkbookToRenderPlan,
  LayoutCursor,
  RenderPlanBuilder,
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
