# 05. Style Registry

## Goal

Tao style registry tap trung. Block/cell chi tham chieu style bang token string, khong inline style tuy tien trong public API.

## Public Shape

```ts
type StyleRegistry = Record<string, CellStyleDefinition>;

interface StyleReference {
  style?: string;
}
```

`CellStyleDefinition` la domain style type cua engine, khong phai ExcelJS style type.

## Rules

- Style token phai ton tai trong `workbook.styles`.
- Style mapping sang ExcelJS nam trong `src/adapters/exceljs`.
- Style khong anh huong layout unless co field explicit da duoc engine support.
- Style errors phai xuat hien truoc khi ghi file.

## Implementation Checklist

- Dinh nghia domain style subset: font, fill, border, alignment, number format.
- Validate style registry names unique theo object key va khong rong.
- Validate moi block/cell style reference.
- Map domain style sang ExcelJS style trong adapter.

## Acceptance

- Reference style token khong ton tai bi reject.
- Public API khong export ExcelJS style type.
- Integration test doc lai `.xlsx` thay style da duoc apply.

