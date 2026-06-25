# Optimize Type-Safety & Reduce Runtime Validation

Mục tiêu: Đẩy mạnh type-safety (Generic Types) ở tầng compile-time của TypeScript và loại bỏ các bước validation cấu trúc bằng `zod` dư thừa ở runtime, chỉ giữ lại các validation về mặt logic/ngữ nghĩa (semantic validation).

## Proposed Changes

### 1. Refactor `types.ts` (Tăng cường Generics, loại bỏ hacky types)
- Xóa bỏ kiểu nội suy rườm rà làm mất Generic như `Extract<Block, { type: 'table' }>['columns'][number]`. Thay vào đó, định nghĩa trực tiếp `TableColumnNode<Row>` và `TableLeafColumn<Row>` với generic `Row`.
- Truyền Generic `Row` xuyên suốt từ `TableBlock<Row>` -> `TableGroup<Row>` -> `TableSectionRow<Row>` -> `TableSectionCell<Row>` -> `TableSectionCellAccessor<Row>`.
- Chỉnh sửa `id?: keyof Row` trong `TableColumn` để TypeScript có thể tự động suggest và kiểm tra lỗi gõ sai tên cột (hiện tại đã có, nhưng sẽ làm chặt hơn, cấm truyền linh tinh).
- Sửa lại các chỗ dùng `unknown` thành Generic linh hoạt hơn (hoặc dùng type parameter mặc định an toàn).

### 2. Dọn dẹp `validation.ts` (Gỡ bỏ Zod cho cấu trúc)
- Vì TypeScript đã đảm bảo cấu trúc JSON khi dev code, việc dùng `zod` để lặp qua từng phần tử mảng (cells, rows, columns, blocks) để kiểm tra `typeof width === 'number'` là cực kỳ tốn tài nguyên và dư thừa.
- **Tháo gỡ hoàn toàn schema cấu trúc của Zod**.
- **Chỉ giữ lại các Semantic Validation quan trọng:**
  - Kiểm tra trùng lặp `sheet.id` và `sheet.name`.
  - Kiểm tra ký tự hợp lệ của tên Sheet (Excel không cho phép `\ / ? * [ ]`).
  - Kiểm tra độ dài tên Sheet (tối đa 31 ký tự).
  - Kiểm tra tính hợp lệ của Reference Style (VD: cột khai báo `style: 'header-style'` nhưng trong workbook registry không có style đó).
  - Kiểm tra `headerRowHeights` không vượt quá độ sâu thực tế của cột header.

### 3. Tận dụng TypeScript `@deprecated`
- Thay vì dùng Runtime Validation để chửi dev khi dùng các field cũ (như `key`, `columnKey`), ta sẽ khai báo thẳng vào `types.ts` với JSDoc `@deprecated Do not use, use id instead` và type là `never`. TypeScript sẽ gạch ngang và báo lỗi đỏ ngay lúc dev đang gõ code.

## Lợi ích
- **Performance tăng vọt:** Bỏ qua hàng ngàn bước check Zod đệ quy tốn CPU, thời gian generate Excel sẽ nhanh hơn rất nhiều.
- **Dev Experience (DX) tốt hơn:** Báo lỗi ngay lập tức trên IDE bằng Type Checker thay vì chờ lúc chạy code mới văng lỗi Zod.
