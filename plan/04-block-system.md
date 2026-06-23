# 04. Block System

## Goal

Dinh nghia block union va compiler registry de moi loai block co compiler rieng, render theo thu tu trong sheet va cap nhat layout cursor on dinh.

## Block Union

Phase nay tao union public:

```ts
type Block =
  | TitleBlock
  | TextBlock
  | SpacerBlock
  | GridBlock
  | TableBlock;
```

Moi block dung `type` lam discriminator. Unknown block type phai bi reject trong validation/compile.

## Compiler Registry

Registry noi bo map:

- `title` -> title compiler.
- `text` -> text compiler.
- `spacer` -> spacer compiler.
- `grid` -> grid compiler.
- `table` -> table compiler.

Compiler signature nen giu on dinh:

```ts
compileBlock(block, context, cursor, builder): void | Promise<void>
```

## Implementation Checklist

- Tao `SheetContext` gom workbook context, sheet id/name, style registry, variable context sau nay.
- Tao registry va default compiler skeleton.
- Skeleton khong duoc sinh fake data; neu block chua duoc support thi throw error ro rang.
- Them test block order va cursor advance cho block co compiler co ban.

## Acceptance

- Unknown block type bi reject ro rang.
- Blocks render dung thu tu khai bao.
- Layout cursor tang dong on dinh sau moi block.

