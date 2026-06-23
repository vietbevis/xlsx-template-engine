# 12. Formula Engine

## Goal

Tao formula abstraction cua engine, khong yeu cau user truyen truc tiep ExcelJS object. Formula compile thanh Excel-compatible formula string trong RenderPlan.

## Formula Model

```ts
type FormulaDefinition =
  | { type: "raw"; expression: string }
  | { type: "literal"; value: string | number | boolean | null }
  | { type: "ref"; key: string }
  | { type: "range"; startKey: string; endKey: string }
  | { type: "sum"; range?: { startKey: string; endKey: string }; values?: FormulaDefinition[] }
  | { type: "round"; value: FormulaDefinition; digits: number }
  | { type: "if"; condition: FormulaDefinition; whenTrue: FormulaDefinition; whenFalse: FormulaDefinition }
  | { type: "call"; name: string; args: FormulaDefinition[] }
  | { type: "binary"; operator: FormulaBinaryOperator; left: FormulaDefinition; right: FormulaDefinition };
```

Phase nay chi can local formula. Cross-sheet ref xu ly o phase 14.

Public formula API khong dung A1/B2, row/column, hoac sheetName. User khai bao key, compiler resolve key thanh dia chi Excel sau khi layout da duoc tinh.

Grid cell co the khai bao `key`. Table leaf column `key` dong thoi la data accessor va formula key trong tung row.

## Formula Context

- Grid: tao key registry theo block, vi grid la layout huu han.
- Table: tao map `key -> columnOffset` mot lan theo leaf columns; moi row dung lightweight context resolve key theo row hien tai. Khong luu key registry cho tung row, de phase streaming/large dataset khong bi giu hang tram nghin dong trong RAM.
- Phase 14 se mo rong context bang `sheetId + key`, khong dung sheetName.

## Rules

- Formula output khong co dau `=` trong domain model; adapter/compiler them khi can.
- Raw formula la escape hatch co kiem soat, validate khong rong va khong bat dau bang `=`.
- Formula cell khong render thanh plain string.
- Formula lồng nhau compile theo AST, vi du `IF(ROUND(SUM(...),0)>20, ..., ...)`.

## Implementation Checklist

- [x] Tao compiler formula domain sang render formula string.
- [x] Ho tro `raw`, `literal`, `ref`, `range`, `sum`, `round`, `if`, `call`, `binary`.
- [x] Tich hop formula vao cell value model.
- [x] Them read-back test xac nhan cell co formula.

## Acceptance

- [x] Formula render ra Excel dung dang `=SUM(...)`.
- [x] Formula cell khong bi coi la string thuong.
- [x] Public API khong lo ExcelJS formula object.
- [x] Public API khong yeu cau user biet A1/row/column.
- [x] Nested formula nhu IF/ROUND/SUM compile duoc tu key.
