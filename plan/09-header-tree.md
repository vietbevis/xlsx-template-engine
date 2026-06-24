# 09. Header Tree Engine

## Goal

Ho tro nested columns va compile Header Tree thanh Header Matrix, Column Schema va merge ranges.

## Column Model

```ts
interface TableColumn<Row> {
  title: string;
  id?: keyof Row;
  accessor?: (row: Row) => CellValue;
  children?: TableColumn<Row>[];
  width?: number;
  style?: string;
}
```

## Compile Flow

```text
Header Tree
  -> normalize/validate
  -> calculate depth and leaf count
  -> Header Matrix
  -> Column Schema
  -> RenderPlan cells and merges
```

## Rules

- Parent column co `children` khong duoc co `id`/`accessor` dung cho data mapping.
- Leaf column bat buoc co `id` hoac `accessor`.
- Engine tu tinh rowSpan/colSpan; user khong khai bao span cho header.
- Leaf order giu dung thu tu input.

## Implementation Checklist

- Tao helper tinh max depth.
- Tao helper flatten leaf columns.
- Tao header matrix voi cell title, row, column, rowSpan, colSpan.
- Reuse merge engine khi phase 10 san sang.

## Acceptance

- Nested header nhieu level render dung merge ngang/doc.
- Column schema leaf order on dinh.
- Parent co id hoac leaf thieu id/accessor bi reject.
