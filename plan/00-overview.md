# 00. Overview

## Muc tieu

Xay dung mot thu vien TypeScript chay tren server de sinh file Excel theo huong Report Engine hoan chinh, khong phai JSON-to-Excel wrapper va khong de nguoi dung phu thuoc truc tiep vao ExcelJS API.

Engine phai giu cac nguyen tac cot loi:

- Header first: header/columns la nguon su that duy nhat cua cau truc bao cao.
- Block based: Workbook -> Sheet -> Block -> Render Plan -> Renderer.
- Streaming first: renderer chinh phai huong toi du lieu lon va commit row tuan tu.
- Adapter isolation: ExcelJS chi nam o tang adapter/render cuoi cung.
- Stable sheet id: sheet dung `id`; moi tham chieu lien sheet dung `sheetId`.

## Architecture Map

Luồng xử lý chuẩn:

```text
WorkbookDefinition
  -> validation
  -> Layout Engine
  -> RenderPlan
  -> Streaming Renderer
  -> ExcelJS Adapter
  -> .xlsx
```

Public API không render trực tiếp từ block xuống ExcelJS. Mọi block compiler chỉ sinh dữ liệu trung gian vào `RenderPlan`.

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
- Không export ExcelJS type qua public API.
- Không dùng `sheet.name` để tham chiếu nội bộ; chỉ dùng `sheet.id`.
- Mỗi phase phải có route API xuất file Excel tương ứng trong `../xlsx-test-app` trước khi coi là xong.
- Khi hoàn thành một phase, phải bổ sung workbook `.xlsx` đầy đủ vào `../xlsx-test-app/src/routes`/`../xlsx-test-app/src/reports` để mở kiểm tra trực quan. `../xlsx-test-app/tests` chỉ gọi route API thật, tải file về `../xlsx-test-app/generated/`, rồi đọc lại workbook để kiểm tra tối thiểu.
- Sau mỗi phase, ghi memory note ngắn theo `plan/MEMORY_POLICY.md`.
