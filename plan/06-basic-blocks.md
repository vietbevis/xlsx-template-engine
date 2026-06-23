# 06. Basic Blocks

## Goal

Implement ba block co ban: title, text, spacer. Day la nen cho report layout doc tu tren xuong duoi.

## Block Shapes

```ts
interface TitleBlock {
  type: "title";
  text: string;
  style?: string;
  height?: number;
}

interface TextBlock {
  type: "text";
  text: string;
  style?: string;
  height?: number;
}

interface SpacerBlock {
  type: "spacer";
  rows?: number;
}
```

## Rules

- `text` co the duoc variable engine xu ly o phase 11; truoc do render raw string.
- `rows` cua spacer mac dinh la 1 neu khong khai bao.
- Height chi set khi user khai bao.
- Khong expose A1/B2 trong public API.

## Implementation Checklist

- Title compiler ghi mot cell tai cursor column mac dinh va advance row.
- Text compiler tuong tu title nhung khong gan semantic title ngoai style token.
- Spacer compiler advance so dong, khong can tao cell value.
- Validate `rows > 0`, `height > 0` neu co.

## Acceptance

- Title/text/spacer tao dung so dong.
- Spacer khong tao data ngoai row trong can thiet.
- Output `.xlsx` doc lai duoc bang ExcelJS.

