# 08. Table Block

## Goal

Implement table data block theo header-first. Columns bat buoc va data chi cung cap gia tri, khong duoc thay doi cau truc report.

## Block Shape

```ts
interface TableBlock<Row = Record<string, unknown>> {
  type: "table";
  columns: TableColumn<Row>[];
  data: Row[] | AsyncIterable<Row>;
  headerStyle?: string;
  bodyStyle?: string;
}

interface TableColumn<Row> {
  title: string;
  id?: keyof Row;
  accessor?: (row: Row) => CellValue;
  width?: number;
  style?: string;
}
```

## Rules

- `columns` bat buoc va khong duoc rong.
- Khong co `autoGenerateColumns`.
- Extra fields trong data bi ignore.
- Moi leaf column can `id` hoac `accessor`.
- Array data support truoc; AsyncIterable hoan thien trong phase 15.

## Implementation Checklist

- Validate table columns.
- Render header row co style token.
- Render body rows bang `id` hoac `accessor`.
- Khong inspect ids/fields cua data de tao column.
- Chuan bi `data` adapter de phase 15 streaming khong phai doi public API.

## Acceptance

- Table thieu columns bi reject.
- Extra field trong data khong tao them cot.
- Header/body render dung theo column order.
