# 02. Workbook And Sheet Foundation

## Goal

Dinh nghia nen tang `WorkbookDefinition`, metadata, sheet registry va validation. Phase nay chot `id` la dinh danh on dinh cua sheet.

## Interfaces

```ts
interface WorkbookMetadata {
  title?: string;
  author?: string;
  company?: string;
  subject?: string;
  keywords?: string[];
}

interface WorkbookDefinition {
  metadata?: WorkbookMetadata;
  styles?: StyleRegistry;
  sheets: SheetDefinition[];
}

interface SheetDefinition {
  id: string;
  name: string;
  blocks: Block[];
}
```

## Validation Rules

- Workbook renderable phai co it nhat mot sheet.
- `sheet.id` bat buoc, khong rong, unique trong workbook.
- `sheet.name` bat buoc, khong rong, unique trong workbook.
- `sheet.name` phai tuan thu gioi han Excel: toi da 31 ky tu va khong chua `: \ / ? * [ ]`.
- Khong co API nao cho phep cross-sheet reference bang `sheet.name`.
- Sheet registry noi bo map `sheetId -> sheet definition/render sheet`.

## Implementation Checklist

- Tao validator rieng cho workbook/sheet.
- Tao error class hoac error helper co message ro rang.
- Map metadata sang render plan metadata, chua can ghi ra Excel ngay neu renderer chua hoan thien.
- Them tests cho duplicate id, duplicate name, invalid name, empty sheets.

## Acceptance

- Duplicate `sheet.id` bi reject.
- Duplicate hoac invalid `sheet.name` bi reject.
- Workbook metadata di qua compile/render pipeline ma khong mat thong tin.

