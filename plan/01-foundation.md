# 01. Foundation

## Goal

Chuyen repo tu template filler hien tai sang nen tang thu vien report engine co public API ro rang, co build/test/typecheck, va co source boundary de cac phase sau khong phai doi lai kien truc.

## Key Changes

- Cap nhat `package.json`:
  - `build`: compile TypeScript ra `dist`.
  - `typecheck`: chay TypeScript no emit.
  - `test`: chay bo test toi thieu.
- Thiet lap source structure:
  - `src/core`: public domain types, validators.
  - `src/compiler`: layout compiler and render-plan builder.
  - `src/renderer`: renderer facade and output APIs.
  - `src/adapters/exceljs`: ExcelJS-specific implementation.
  - `src/index.ts`: public exports only.
- Loai bo public API cu dang template filler hoac de lai sau mot facade migration neu can, nhung khong de API cu dinh huong kien truc moi.

## Public API Draft

```ts
const workbook = defineWorkbook({
  metadata: { title: "Report" },
  sheets: [
    {
      id: "summary",
      name: "Summary",
      blocks: [],
    },
  ],
});

await renderWorkbook(workbook).writeFile("report.xlsx");
```

`renderWorkbook()` tra ve writer facade, khong tra ve ExcelJS workbook.

## Implementation Checklist

- Tao domain type toi thieu cho workbook, sheet, block placeholder.
- Tao `defineWorkbook()` de normalize va validate shape co ban.
- Tao `renderWorkbook()` facade nhan workbook definition.
- Tao renderer stub co `writeFile`, `writeBuffer`, `writeStream` nhung chi implement muc an toan toi thieu khi phase 2/3 san sang.
- Dam bao `src/index.ts` chi export API du kien dung lau dai.

## Acceptance

- `npm run build` tao duoc `dist`.
- `npm run typecheck` pass.
- Public API khong export truc tiep ExcelJS type.
- Smoke test import package, tao workbook definition hop le va goi validation.

