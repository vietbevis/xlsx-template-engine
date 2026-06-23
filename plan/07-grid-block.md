# 07. Grid Block

## Goal

Implement grid layout tu do cho form/report sections co cau truc bang co dinh, khong phai data table streaming.

## Block Shape

```ts
interface GridBlock {
  type: "grid";
  rows: GridRow[];
}

interface GridRow {
  height?: number;
  cells: GridCell[];
}

interface GridCell {
  value?: CellValue;
  style?: string;
  colSpan?: number;
  rowSpan?: number;
  width?: number;
}
```

## Rules

- Grid cell layout theo vi tri tuong doi, bat dau tai cursor.
- `rowSpan` va `colSpan` mac dinh la 1.
- Merge cua grid phai di qua merge validation chung khi phase 10 co san.
- Width/height chi set khi khai bao explicit.
- Grid khong scan data va khong sinh cot dong.

## Implementation Checklist

- Compile rows/cells thanh render cells va merge ranges.
- Validate span la so nguyen duong.
- Detect overlap trong grid local ngay ca truoc phase merge engine day du.
- Advance cursor theo so row cua grid.

## Acceptance

- Grid cell render dung vi tri tuong doi.
- Merge overlap bi reject.
- Width/height chi xuat hien trong render plan khi user khai bao.

