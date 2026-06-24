# 16. Release Hardening

## Goal

Hoan thien docs, examples, tests va performance notes de thu vien san sang dung thuc te.

## Examples

Them examples chay bang `tsx`:

- Simple workbook.
- Styled report.
- Table with nested header.
- Multi-sheet summary.
- Streaming large table.

## Documentation

- README mo ta engine moi, khong con mo ta nhu template filler.
- API docs cho `defineWorkbook`, `renderWorkbook`, blocks, styles, formulas.
- Unsupported features ro rang:
  - no auto-generated columns
  - no auto width
  - no auto height
  - array table data is supported; `AsyncIterable` data is not implemented yet
  - `writeBuffer` uses the non-streaming workbook path
  - current `RenderPlan` still materializes rows before the adapter writes them
  - ExcelJS primitives are reused directly for workbook values/styles where useful

## Tests

- `../xlsx-test-app` la noi bat buoc chua route API xuat `.xlsx` cho moi phase da hoan thanh; test phai goi route that va luu file de mo kiem tra.
- Unit tests cho validators, header compiler, merge engine, variable resolver, dependency graph.
- Integration tests sinh `.xlsx`, doc lai bang ExcelJS de assert sheet/cell/merge/style/formula.
- Streaming tests voi AsyncIterable.
- Type/public export tests neu co tooling phu hop.

## Acceptance

- Examples chay duoc bang `tsx`.
- Build/typecheck/test pass.
- README va plan dong bo voi behavior thuc te.
