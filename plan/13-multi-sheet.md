# 13. Multi Sheet

## Goal

Ho tro render nhieu sheet trong mot workbook, voi registry noi bo dua tren `sheetId`.

## Rules

- Render order mac dinh theo thu tu `workbook.sheets`.
- Sheet registry map `sheet.id -> sheet.name`.
- Duplicate id/name da duoc validate o phase 2.
- Public cross-sheet references tiep tuc dung `sheetId`; khong them API tham chieu bang name.

## Implementation Checklist

- Cap nhat compiler de lap qua tat ca sheets.
- Tao render plan sheet cho moi sheet.
- Adapter tao worksheet theo sheet order trong render plan.
- Chuan bi hook dependency graph cho phase 14 nhung chua doi order neu chua can.

## Acceptance

- Workbook nhieu sheet ghi file thanh cong.
- Doc lai file thay dung sheet names va order.
- Public API van dung `id`/`sheetId`.

