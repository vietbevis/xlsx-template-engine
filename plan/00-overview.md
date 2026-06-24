# 00. Overview

## Muc tieu

Xay dung mot thu vien TypeScript chay tren server de sinh file Excel theo huong Report Engine hoan chinh, khong phai JSON-to-Excel wrapper. ExcelJS la loi workbook/cell/style cua thu vien; engine chi them report DSL, layout block, semantic id, formula resolver va streaming workflow.

Engine phai giu cac nguyen tac cot loi:

- Header first: header/columns la nguon su that duy nhat cua cau truc bao cao.
- Block based: Workbook -> Sheet -> Block -> Render Plan -> Renderer.
- Streaming first: renderer chinh phai huong toi du lieu lon va commit row tuan tu.
- ExcelJS core-first: dung truc tiep type/primitive cua ExcelJS cho workbook value, style, border, fill, font va alignment thay vi dinh nghia lai.
- Stable sheet id: sheet dung `id`; moi tham chieu lien sheet dung `sheetId`.

## Architecture Map

Luồng xử lý chuẩn:

```text
WorkbookDefinition
  -> validation
  -> Layout Engine
  -> ExcelJS-backed RenderPlan
  -> Streaming Renderer
  -> .xlsx
```

Public API không bắt người dùng tự tính A1/row/column. Block compiler vẫn sinh dữ liệu trung gian vào `RenderPlan`, nhưng các primitive cell/style/value trong plan lấy từ ExcelJS.

## Phase Map

- Phase 1: Foundation, package scripts, public entrypoint, source boundaries.
- Phase 2: Workbook/Sheet foundation, metadata, validation.
- Phase 3: Render Plan architecture.
- Phase 4: Block system and compiler registry.
- Phase 5: Style registry.
- Phase 6: Title/Text/Spacer blocks.
- Phase 7: Grid block.
- Phase 8: Table block.
- Phase 9: Header tree compiler.
- Phase 10: Merge engine.
- Phase 11: Variable engine.
- Phase 12: Formula engine.
- Phase 13: Multi-sheet support.
- Phase 14: Cross-sheet formula/link/dependency graph.
- Phase 15: Streaming renderer.
- Phase 16: Release hardening.

## Development Rules

- Không thêm cột từ data, kể cả khi data có field dư.
- Không scan toàn bộ data để tính width/height.
- Không định nghĩa lại type ExcelJS đã có; public type có thể alias về ExcelJS khi đó là primitive của workbook/cell/style.
- Không dùng `sheet.name` để tham chiếu nội bộ; chỉ dùng `sheet.id`.
- Mỗi phase phải có route API xuất file Excel tương ứng trong `../xlsx-test-app` trước khi coi là xong.
- Khi hoàn thành một phase, phải bổ sung workbook `.xlsx` đầy đủ vào `../xlsx-test-app/src/routes`/`../xlsx-test-app/src/reports` để mở kiểm tra trực quan. `../xlsx-test-app/tests` chỉ gọi route API thật, tải file về `../xlsx-test-app/generated/`, rồi đọc lại workbook để kiểm tra tối thiểu.
- Sau mỗi phase, ghi memory note ngắn theo `plan/MEMORY_POLICY.md`.
