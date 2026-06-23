# 15. Streaming Renderer

## Goal

Dung ExcelJS Streaming Writer lam renderer chinh cho `writeFile` va `writeStream`, de engine co the xu ly dataset lon ma khong giu toan bo workbook trong RAM.

## Output APIs

```ts
renderWorkbook(workbook, context).writeFile(path)
renderWorkbook(workbook, context).writeStream(stream)
renderWorkbook(workbook, context).writeBuffer()
```

Policy:

- `writeFile` va `writeStream` uu tien streaming writer.
- `writeBuffer` co the dung non-streaming path neu ExcelJS streaming khong tra buffer tien loi; docs phai ghi ro gioi han.

## Data Sources

Support theo thu tu:

- Array.
- AsyncIterable.
- Node Stream neu co adapter ro rang.
- Database cursor thong qua AsyncIterable.

## Streaming Rules

- Row duoc write va commit tuan tu.
- Khong scan data de tinh width/height.
- Table compiler khong materialize toan bo AsyncIterable vao memory.
- RenderPlan cho large table can cho phep lazy row producer thay vi giu tat ca rows.

## Acceptance

- Render duoc 10k va 100k rows bang streaming path.
- Memory usage khong tang theo kieu giu toan workbook/data trong RAM.
- Row duoc commit tuan tu.

