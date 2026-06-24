# 13. Multi Sheet

## Goal

Ho tro render nhieu sheet trong mot workbook, voi registry noi bo dua tren `sheetId`.

## Rules

- Render order mac dinh theo thu tu `workbook.sheets`.
- Sheet registry map `sheet.id -> sheet.name`.
- Duplicate id/name da duoc validate o phase 2.
- Public cross-sheet references tiep tuc dung `sheetId`; khong them API tham chieu bang name.

## Implementation Checklist

- [x] Cap nhat compiler de lap qua tat ca sheets.
- [x] Tao render plan sheet cho moi sheet.
- [x] Adapter tao worksheet theo sheet order trong render plan.
- [x] Chuan bi hook dependency graph cho phase 14 nhung chua doi order neu chua can.

## Acceptance

- [x] Workbook nhieu sheet ghi file thanh cong.
- [x] Doc lai file thay dung sheet names va order.
- [x] Public API van dung `id`/`sheetId`.

## Notes

- Phase 13 giu render order theo `workbook.sheets`; dependency graph phase 14 se co collector rieng va chi doi order neu co nhu cau ro.
- Local formula id registry khong global qua sheet. Cross-sheet formula phase 14 se mo rong bang `sheetId + id`, khong dung sheet name.
