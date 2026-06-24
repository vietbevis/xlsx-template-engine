# Sequential Feature Expansion Plan

## Summary

- Triển khai lần lượt các nhóm mở rộng production-facing, dùng dialect hiện tại của engine là `id/startId/endId/columnId`, không quay lại `key`.
- Bỏ qua `ImageBlock` và `ChartBlock` theo quyết định mới của bạn.
- Ưu tiên P0/P1 trước: error hierarchy, freeze pane, table footer/summary, conditional style, formula helpers/types.
- Bỏ qua `pageSetup`, print area, repeat print rows và các setting in ấn vì engine chạy server-side, không cần tối ưu luồng in trực tiếp.
- Sau khi hoàn tất P0/P1, thì đặc biệt triển khai chi tiết P2
- Mỗi cụm feature phải có unit/integration test, type-safety test nếu đổi public API, và workbook artifact inspectable khi output user-facing.

## Phase 1 — P0 Foundation

- Error hierarchy:
  - Giữ `ReportEngineError` base class.
  - Thêm `ValidationError`, `CompileError`, `FormulaError`, `RenderError`.
  - Validation errors nhận `{ message, path }`; compile/formula errors nhận `{ sheetId, blockIndex?, id? }`.
  - Không parse message trong tests nữa; assert bằng `instanceof` và structured fields.
- Freeze pane:
  - Thêm `SheetDefinition.freezePane?: { rows?: number; columns?: number }`.
  - Thêm `RenderPlanSheet.views` hoặc field tương đương.
  - ExcelJS adapter map sang `views: [{ state: 'frozen', xSplit, ySplit }]` khi tạo worksheet.
  - Validate rows/columns là positive integer nếu có.
- Page setup / print:
  - Không triển khai trong rollout này.
  - Không thêm `PageSetupBlock`, `printArea`, `repeatRows`, margin/fit-to-page hoặc mapping ExcelJS `pageSetup`.

## Phase 2 — P1 Table and Style

- Conditional style:
  - Add `styleResolver?: (value: CellValue | undefined) => StyleValue | undefined` to `GridCell` and `TableSectionCell`.
  - Add `styleResolver?: (value: CellValue | undefined, row: Row, rowIndex: number) => StyleValue | undefined` to `TableColumn`.
  - Resolve style during compile after value/formula result is known; precedence: `defaultStyle < style/styleResolver result < inlineStyle`.
  - Validate returned style tokens against workbook styles.
- Table footer rows:
  - Add `footerRows?: readonly TableFooterRow<Row>[]`.
  - `TableFooterRow` reuses `TableSectionCell<Row>[]`, `style`, `height`.
  - Render footer after all data rows and section rows.
  - Footer formula scope supports `allRows`; `currentRows` in footer means all table data rows.
- Band/row style:
  - Add table-level `evenRowStyle?`, `oddRowStyle?`; column-level overrides allowed.
  - Apply to body data rows only, not title/header/footer rows.
- Column/row visibility:
  - Add `TableColumn.hidden?: boolean`.
  - Hidden columns still exist in formula id map and data lookup, but adapter marks column hidden.
  - Add `TableDataItem` row visibility via section/data row metadata: `{ hidden?: boolean }` for section rows and a table-level `rowHidden?: (row, index) => boolean` for data rows.
  - Hidden rows still exist for formulas but adapter marks row hidden.
- Summary shorthand:
  - Add `TableColumn.summary?: 'sum' | 'count' | 'average' | FormulaDefinition`.
  - Add `summaryStyle?: StyleValue`.
  - Compiler creates one footer row when at least one leaf column has summary and no explicit summary footer conflict.
  - Summary formulas target the column id across all visible and hidden data rows; hidden rows remain included by default.

## Phase 3 — P2 Formula Engine

[Chi tiết plan cho phase này trong file](./plan-formula-builder.md)

## Phase 4 — Block System Without Image/Chart

- Divider block:
  - Add `{ type: 'divider'; style?: StyleValue; rows?: number }`.
  - Emits one or more blank rows with style applied across current known table width if available, otherwise first column.
- Repeat block:
  - Add generic `RepeatBlock<T>`.
  - Shape: `{ type: 'repeat'; data: readonly T[]; block: (item: T, index: number) => Block | readonly Block[] }`.
  - Compiler expands repeat blocks in order before normal block compilation for that sheet.
  - Validation checks `data` array and `block` function; compile errors include parent repeat block path.
- Explicitly skip:
  - Do not implement `ImageBlock`.
  - Do not implement `ChartBlock`.
  - Docs should state both are out of scope for this rollout.

## Phase 5 — Architecture and DX

- Plugin system:
  - Add `plugins?: WorkbookPlugin[]` to `renderWorkbook` and compile options.
  - Hooks: `onBeforeCompile`, `onAfterCompile`, `onBeforeRender`, `onAfterRenderSheet`.
  - Hooks may transform workbook/render plan where declared; validation reruns after `onBeforeCompile`.
- Async block compiler:
  - Add `compileWorkbookToRenderPlanAsync`.
  - Add `renderWorkbookAsync` or make renderer construction async with a new API; keep existing sync API unchanged.
  - `BlockCompiler` advanced type becomes `void | Promise<void>` only for async compile path.
- Diagnostics:
  - Add `compileWorkbook(workbook, options): { renderPlan, diagnostics }`.
  - Warnings include missing column width, empty sheet, missing formula result cache, large merge.
  - Existing `compileWorkbookToRenderPlan` remains throw-on-error and returns only `RenderPlan`.
- Incremental compile:
  - Add `createWorkbookCompiler(workbook)` in advanced API.
  - Supports `compileSheet(sheetId)` and dependency-aware sheet id validation.
  - Do not parallelize yet; just expose stable per-sheet compile boundary.
- Debug utilities:
  - Add `diffRenderPlans(left, right)` in `src/debug.ts`.
  - Add a debug renderer that returns/prints ASCII-like sheet summaries from RenderPlan, not a full Excel writer.
- Schema export:
  - Add `generateWorkbookSchema()` in `src/schema.ts`.
  - Generate static JSON Schema from maintained schema builders, not from TypeScript reflection.

## Test Plan

- Run after each phase: `npm run typecheck`, `npm run typecheck:tests`, `npm test`, `npm run lint`, `npm run build`, `git diff --check`.
- Add type-safety assertions for all new public fields and old `key` examples staying rejected.
- Add ExcelJS read-back tests for freeze panes, hidden rows/columns, named ranges, formula cache, footer rows, summary formulas.
- Add adapter tests for both streaming-compatible features and non-streaming-only behavior where relevant.
- Add generated workbook artifacts under `output/` for P0/P1 user-facing output, especially page setup/freeze/table footer examples.
- Update `src/examples/overtime-report.ts` to use freeze panes, footer/summary shorthand, and conditional styles.

## Assumptions

- All new APIs use `id`, not `key`.
- `ImageBlock` and `ChartBlock` are intentionally skipped.
- Existing sync APIs remain backward compatible except for additive type/API changes.
- True lazy `AsyncIterable` table rendering remains separate from these features unless the async compiler phase explicitly needs it.
