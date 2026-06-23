# 10. Merge Engine

## Goal

Tao merge model chung cho grid, table header va cac feature sau nay. Merge engine phai validate truoc khi renderer ghi Excel.

## Merge Model

```ts
interface MergeRange {
  sheetId: string;
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}
```

## Validation Rules

- Row/column one-based va la so nguyen duong.
- `endRow >= startRow`, `endColumn >= startColumn`.
- Merge one-cell co the bo qua hoac reject theo mot policy co dinh; mac dinh bo qua.
- Khong cho overlap merge ranges trong cung sheet.
- Khong merge vuot bounds da compile neu sheet bounds da biet.

## Implementation Checklist

- Tao utility normalize merge range.
- Tao overlap detector theo sheet.
- Tich hop vao RenderPlanBuilder.
- Adapter chi apply merge da validate.

## Acceptance

- Horizontal, vertical, mixed merge deu pass.
- Overlap merge fail truoc khi ghi Excel.
- Header merge khong can user khai bao rowSpan/colSpan.

