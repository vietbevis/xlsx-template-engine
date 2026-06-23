Excel Report Engine Specification

1. Mục tiêu

Xây dựng một thư viện TypeScript chạy trên môi trường server để sinh file Excel.

Thư viện sử dụng ExcelJS làm tầng render cuối cùng nhưng không phụ thuộc trực tiếp vào API của ExcelJS.

Mục tiêu là xây dựng một Report Engine hoàn chỉnh, không phải một JSON-to-Excel wrapper.

⸻

2. Nguyên tắc thiết kế

Header First Architecture

Header là nguồn sự thật duy nhất.

Data chỉ cung cấp giá trị.

Data không được phép thay đổi cấu trúc báo cáo.

Đúng

columns = [
  {
    title: "Toán",
    key: "math"
  }
]
data = [
  {
    math: 10
  }
]

Sai

data = [
  {
    math: 10,
    chemistry: 8
  }
]

và tự động sinh thêm cột Hóa.

Không hỗ trợ auto generate column từ data.

⸻

Block Based Layout

Không làm việc trực tiếp với A1, B2, C3.

Workbook được cấu thành từ các Sheet.

Sheet được cấu thành từ các Block.

Workbook
 └── Sheet
      └── Block[]

⸻

Streaming First Architecture

Renderer chính sử dụng Streaming Writer.

Mục tiêu:

10.000 rows
100.000 rows
500.000 rows
1.000.000+ rows

Không thiết kế theo hướng giữ toàn bộ workbook trong RAM.

⸻

3. Workbook

Workbook

Là root object.

Chứa:

metadata
sheets

⸻

Workbook Metadata

interface WorkbookMetadata {
  title?: string
  author?: string
  company?: string
  subject?: string
  keywords?: string[]
}

⸻

4. Sheet

Sheet Definition

interface SheetDefinition {
  id: string
  name: string
}

⸻

Quy tắc

id

Bắt buộc.

Ổn định.

Không thay đổi.

Dùng nội bộ engine.

name

Tên hiển thị trên Excel.

Có thể thay đổi.

Không dùng để tham chiếu.

⸻

Ví dụ

{
  id: "department_cntt",
  name: "CNTT"
}

⸻

5. Cross Sheet Reference

Mọi tham chiếu liên sheet đều dùng:

sheetId

không dùng:

sheetName

⸻

Ví dụ

{
  ref: {
    sheetId: "department_cntt",
    cell: "F10"
  }
}

Engine tự map:

='CNTT'!F10

⸻

6. Sheet Dependency Graph

Hỗ trợ đồ thị phụ thuộc giữa các sheet.

Ví dụ:

CNTT ----\
KETOAN --- > SUMMARY
NGOAINGU-/

Engine biết:

* sheet nào phụ thuộc sheet nào
* render order
* formula dependency

⸻

7. Block System

Sheet chứa danh sách Block.

interface SheetDefinition {
  id: string
  name: string
  blocks: Block[]
}

⸻

8. Block Types

Title Block

{
  type: "title"
}

⸻

Text Block

{
  type: "text"
}

⸻

Spacer Block

{
  type: "spacer"
}

⸻

Grid Block

{
  type: "grid"
}

Cho phép tạo layout dạng bảng tùy ý.

⸻

Table Block

{
  type: "table"
}

Dùng cho dữ liệu dạng bảng.

⸻

9. Table Engine

Table luôn phải có:

columns

Không hỗ trợ:

autoGenerateColumns

⸻

10. Header Tree Engine

Header được định nghĩa dạng cây.

Ví dụ:

[
  {
    title: "HK1",
    children: [
      {
        title: "Toán",
        key: "math"
      },
      {
        title: "Lý",
        key: "physics"
      }
    ]
  }
]

⸻

Compile

Header Tree

↓

Header Matrix

↓

Column Schema

↓

Render

⸻

11. Merge Engine

Hỗ trợ:

rowSpan
colSpan

theo semantics HTML Table.

⸻

Không bắt người dùng khai báo

rowSpan
colSpan

cho Header.

Engine tự tính từ Header Tree.

⸻

Hỗ trợ

Horizontal Merge

A1:D1

Vertical Merge

A1:A3

Mixed Merge

A1:D3

⸻

12. Grid Engine

Cho phép tạo layout dạng bảng tự do.

Ví dụ:

{
  type: "grid",
  rows: [...]
}

⸻

13. Style Registry

Không sử dụng inline style tùy tiện.

Định nghĩa style tập trung.

styles = {
  reportTitle: {},
  tableHeader: {},
  tableBody: {}
}

Block chỉ tham chiếu:

style: "reportTitle"

⸻

14. Variable Engine

Hỗ trợ template.

Ví dụ:

{{today}}
{{student.name}}
{{students.length}}

⸻

15. Formula Engine

Hỗ trợ formula.

Ví dụ:

{
  formula: {}
}

Render:

=SUM(...)

⸻

16. Cross Sheet Formula Engine

Ví dụ:

{
  formula: {
    ref: {
      sheetId: "department_cntt",
      cell: "F10"
    }
  }
}

Render:

='CNTT'!F10

⸻

17. Sheet Link Engine

Hỗ trợ hyperlink giữa các sheet.

Tham chiếu bằng:

sheetId

Engine tự map sang sheet name.

⸻

18. Multi Sheet

Hỗ trợ nhiều sheet.

Ví dụ:

Summary
CNTT
KETOAN
NGOAINGU

⸻

19. Render Architecture

Workbook
 ↓
Sheet
 ↓
Block
 ↓
Layout Engine
 ↓
Render Plan
 ↓
Streaming Renderer
 ↓
ExcelJS

⸻

20. Layout Engine

Nhiệm vụ:

* compile header tree
* tính merge
* tính formula
* build dependency graph
* build render plan

⸻

21. Render Plan

Render Plan là representation trung gian.

Không render trực tiếp từ Block xuống ExcelJS.

⸻

22. Streaming Renderer

Renderer chính.

Sau khi Render Plan hoàn tất:

Read Row
 ↓
Map Value
 ↓
Write Row
 ↓
Commit Row

⸻

23. Không hỗ trợ Auto Width Engine

Không scan dữ liệu để tính width.

Nếu người dùng không chỉ định:

width

thì không set.

⸻

24. Không hỗ trợ Auto Height Engine

Không scan dữ liệu để tính height.

Nếu người dùng không chỉ định:

height

thì không set.

⸻

25. Data Source

Thiết kế hướng tới:

Array
AsyncIterable
Stream
Database Cursor

để xử lý dữ liệu lớn.

⸻

26. Roadmap

Phase 1

* Workbook
* Workbook Metadata
* Sheet Definition (id + name)
* Workbook Builder
* Workbook Renderer
* ExcelJS Adapter
* writeFile()
* writeBuffer()

⸻

Phase 2

* Sheet Foundation
* Sheet Context
* Sheet Renderer

⸻

Phase 3

* Block System

⸻

Phase 4

* Style Registry

⸻

Phase 5

* Title Block

⸻

Phase 6

* Text Block

⸻

Phase 7

* Spacer Block

⸻

Phase 8

* Grid Block

⸻

Phase 9

* Table Block

⸻

Phase 10

* Header Tree Engine

⸻

Phase 11

* Merge Engine

⸻

Phase 12

* Variable Engine

⸻

Phase 13

* Formula Engine

⸻

Phase 14

* Multi Sheet

⸻

Phase 15

* Cross-Sheet Formula Engine

⸻

Phase 16

* Sheet Link Engine

:::
Có một điểm duy nhất tôi đề xuất điều chỉnh trước khi bắt đầu code sâu hơn:
**Sheet Definition nên có `key` thay vì `id` hoặc hỗ trợ cả hai nhưng chỉ giữ một khái niệm duy nhất.**
Ví dụ:
```ts
interface SheetDefinition {
  key: string
  name: string
}

hoặc

interface SheetDefinition {
  id: string
  name: string
}

Không nên tồn tại đồng thời id và key vì sau này toàn bộ Formula Engine, Dependency Graph, Hyperlink Engine, Named Range Mapping đều sẽ phụ thuộc vào định danh này. Chọn một tên duy nhất từ đầu sẽ tránh rất nhiều refactor về sau. Tôi nghiêng về key hơn vì nó gần với khái niệm “stable identifier” trong các engine báo cáo.