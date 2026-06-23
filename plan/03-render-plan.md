# 03. Render Plan Architecture

## Goal

Tao intermediate representation de tach layout/domain khoi ExcelJS. Sau phase nay, renderer chi lam viec voi `RenderPlan`, khong doc `Block` truc tiep.

## Render Plan Model

Render plan can bieu dien toi thieu:

- Workbook metadata.
- Sheets theo `sheetId`, `sheetName`.
- Row commands theo thu tu streaming.
- Cell commands gom row, column, value, style token, formula/link neu co.
- Merge ranges.
- Column widths va row heights explicit.
- Style registry da validate.

## Layout Cursor

`LayoutCursor` quan ly vi tri trong sheet:

- `row`: dong hien tai, one-based theo Excel semantics.
- `column`: cot bat dau mac dinh, one-based.
- helper advance row/column.

Public block API khong dung A1/B2; compiler co the doi cursor thanh toa do khi tao render plan.

## Implementation Checklist

- Tao `compileWorkbookToRenderPlan(workbook)`.
- Tao `RenderPlanBuilder` de block compiler them row/cell/merge/style.
- Renderer facade goi compiler truoc, sau do truyen plan vao adapter.
- Khong import ExcelJS trong `src/core` va `src/compiler`.

## Acceptance

- Workbook don gian compile duoc thanh render plan.
- Render plan co du thong tin sheet id/name va metadata.
- Test hoac static check xac nhan compiler layer khong import ExcelJS.

