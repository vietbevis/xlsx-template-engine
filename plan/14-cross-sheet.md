# 14. Cross Sheet Features

## Goal

Ho tro dependency graph, cross-sheet formula va sheet hyperlink bang `sheetId`.

## Cross-Sheet Reference

```ts
{
  type: "ref",
  sheetId: "department_cntt",
  id:  "total_hours"
}
```

Compiler map sang:

```text
='CNTT'!F10
```

Sheet name phai duoc quote/escape dung Excel syntax.

Public API van khong dung A1/F10. `F10` chi la dia chi Excel duoc compiler resolve tu `sheetId + id` sau khi layout da biet.

## Dependency Graph

Graph can biet:

- Sheet nao phu thuoc sheet nao.
- Missing sheetId.
- Circular dependency neu dependency anh huong render order.
- Formula/link dependency cho diagnostics.

## Implementation Checklist

- [x] Tao collector de quet formulas trong workbook definition.
- [x] Validate moi `sheetId` ton tai trong sheet registry.
- [x] Tao dependency graph hook; chua doi render order khi chua co nhu cau ro.
- [x] Cross-sheet formula compile bang `sheetId + id`.
- [ ] Implement sheet hyperlink compile bang `sheetId`.

## Acceptance

- [x] `sheetId` duoc map sang quoted sheet name.
- [x] Missing `sheetId` fail ro rang.
- [x] Rename `sheet.name` khong lam hong formula neu `id` giu nguyen.

## Notes

- Workbook-level id registry chi thu thap grid cells co `id`, phu hop voi summary/subtotal cells va khong giu tung row table trong RAM.
- Table formulas tiep tuc dung row-local id context de san sang cho phase 15 streaming.
- Sheet hyperlink la surface rieng vi public block API chua co link definition; giu lai cho phase link/UI sau.
