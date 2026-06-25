# Project Structure

Project hiện dùng cấu trúc phẳng để dễ đọc và dễ maintain. Không còn chia source theo
`core/`, `compiler/`, `adapter/`, `renderer/`.

## Source

- `src/types.ts`: toàn bộ type public của workbook, sheet, block, table, formula, style.
- `src/workbook.ts`: runtime check workbook.
- `src/validation.ts`: validate workbook input.
- `src/compile.ts`: compile workbook thành render plan.
- `src/block-compiler.ts`: render từng block vào render plan builder.
- `src/render-plan.ts`: type của render plan.
- `src/render-plan-builder.ts`: helper build render plan.
- `src/exceljs-workbook.ts`: ghi render plan ra `.xlsx` bằng ExcelJS.
- `src/renderer.ts`: API `renderWorkbook`.
- `src/formula.ts`: helper tạo formula DSL.
- `src/formula-engine.ts`: compile formula DSL thành Excel formula string.
- `src/variable-engine.ts`: interpolate `{{value}}`.
- `src/layout-cursor.ts`: cursor row/column khi compile.
- `src/merge-engine.ts`: validate merge range.
- `src/index.ts`: public exports.

## Helpers

Logic dùng chung đặt trong `src/helpers/`:

- `common.ts`: object guard, positive integer assert, exhaustive assert.
- `grid.ts`: grid occupancy, overlap check, `colSpan: "remaining"`.
- `style.ts`: clone style và metadata.
- `table.ts`: flatten columns, header matrix, table section cell resolver, summary formula.
- `workbook.ts`: named range resolver và đo số cột sheet.
- `utils.ts`: re-export các helper trên để import ngắn.

## Rule

Khi có logic dùng lại từ 2 nơi trở lên, đưa vào `src/helpers/`. File chính chỉ nên giữ
flow đọc được từ trên xuống dưới, hạn chế duplicate helper riêng lẻ trong từng file.
