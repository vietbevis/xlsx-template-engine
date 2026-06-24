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

`CellStyleDefinition` la `Partial<ExcelJS.Style>`. Khong tao lai fill/font/border/alignment/number format rieng cua engine.

## Rules

- Style token phai ton tai trong `workbook.styles`.
- Style registry dung truc tiep ExcelJS style shape; renderer gan style vao cell, khong dich alias rieng.
- Style khong anh huong layout unless co field explicit da duoc engine support.
- Style errors phai xuat hien truoc khi ghi file.

## Implementation Checklist

- Tai su dung ExcelJS style type cho font, fill, border, alignment, number format.
- Validate style registry names unique theo object key va khong rong.
- Validate moi block/cell style reference.
- Renderer gan ExcelJS style truc tiep.

## Acceptance

- Reference style token khong ton tai bi reject.
- Public API khong dinh nghia lai ExcelJS style type.
- Integration test doc lai `.xlsx` thay style da duoc apply.
