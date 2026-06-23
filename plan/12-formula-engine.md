# 12. Formula Engine

## Goal

Tao formula abstraction cua engine, khong yeu cau user truyen truc tiep ExcelJS object. Formula compile thanh Excel-compatible formula string trong RenderPlan.

## Formula Model

```ts
type FormulaDefinition =
  | { type: "raw"; expression: string }
  | { type: "sum"; range: CellRangeReference }
  | { type: "ref"; ref: CellReference };
```

Phase nay chi can local formula. Cross-sheet ref xu ly o phase 14.

## Rules

- Formula output khong co dau `=` trong domain model; adapter/compiler them khi can.
- Raw formula la escape hatch co kiem soat, validate khong rong.
- Formula cell khong render thanh plain string.

## Implementation Checklist

- Tao compiler formula domain sang render formula string.
- Ho tro `raw` va toi thieu mot helper nhu `sum`.
- Tich hop formula vao cell value model.
- Them read-back test xac nhan cell co formula.

## Acceptance

- Formula render ra Excel dung dang `=SUM(...)`.
- Formula cell khong bi coi la string thuong.
- Public API khong lo ExcelJS formula object.

