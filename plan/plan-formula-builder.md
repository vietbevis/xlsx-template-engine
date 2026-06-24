# Phase — Formula Builder & Engine Expansion

## Mục tiêu

Mở rộng formula engine theo ba hướng song song:

1. **Builder API** (`f.*`) — fluent helper giảm verbosity, không thay đổi representation cốt lõi.
2. **Named Ranges** — đặt tên vùng ô để formula và Excel UI tham chiếu bằng tên thay vì địa chỉ.
3. **Typed formula variants** cho các hàm Excel phổ biến còn thiếu, và **formula result cache**.

Ba hướng độc lập với nhau — builder chỉ là syntactic sugar, named range là feature mới trong pipeline, typed variants là mở rộng `FormulaDefinition` union. Thứ tự implement có thể linh hoạt nhưng nên hoàn thành cả ba trước khi đóng phase.

---

## 1. Formula Builder API

### Nguyên tắc thiết kế

Builder trả về đúng `FormulaDefinition` object — không có runtime abstraction riêng, không có class, không có lazy evaluation. User có thể mix builder và object literal tùy ý vì output là cùng kiểu.

```ts
import { f } from 'xlsx-template-engine';

// Builder và object literal hoàn toàn tương đương
f.ref('score');
// === { type: 'ref', id: 'score' }

f.mul(f.ref('hours'), f.ref('rate'));
// === { type: 'binary', operator: '*', left: { type: 'ref', id: 'hours' }, right: { type: 'ref', id: 'rate' } }
```

### File mới: `src/formula.ts`

Export duy nhất là object `f`. Không export từng helper riêng lẻ để tránh namespace pollution.

```ts
export const f = {
  // --- Refs ---
  ref(id: string): RefFormulaDefinition
  xref(sheetId: string, id: string): RefFormulaDefinition

  // --- Range ---
  range(startId: string, endId: string, options?: { sheetId?: string; scope?: FormulaRangeScope }): RangeFormulaDefinition

  // --- Named range ---
  namedRange(name: string): NamedRangeFormulaDefinition

  // --- Literal ---
  // Không cần { type: 'literal', value: 20 } trong binary/if — coerce tự động
  val(value: string | number | boolean | null): LiteralFormulaDefinition

  // --- Raw ---
  raw(expression: string): RawFormulaDefinition

  // --- Aggregation ---
  // Tách rõ hai mode của sum hiện tại thành hai helper khác nhau
  sumRange(startId: string, endId: string, options?: { sheetId?: string; scope?: FormulaRangeScope }): SumFormulaDefinition
  sum(...values: FormulaDefinition[]): SumFormulaDefinition

  // --- Math ---
  add(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition
  sub(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition
  mul(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition
  div(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition
  round(value: FormulaDefinition, digits: number): RoundFormulaDefinition

  // --- Comparison ---
  // right nhận FormulaDefinition hoặc primitive — coerce tự động thành literal
  gt(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition
  gte(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition
  lt(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition
  lte(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition
  eq(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition
  neq(left: FormulaDefinition, right: CoercibleValue): BinaryFormulaDefinition

  // --- Control flow ---
  // whenTrue/whenFalse nhận primitive — coerce tự động
  if(
    condition: FormulaDefinition,
    whenTrue: CoercibleValue,
    whenFalse: CoercibleValue
  ): IfFormulaDefinition

  // --- Typed Excel functions ---
  // Các hàm này compile thành CallFormulaDefinition hoặc typed variant tùy thiết kế bên dưới
  max(...values: FormulaDefinition[]): MaxFormulaDefinition
  min(...values: FormulaDefinition[]): MinFormulaDefinition
  average(startId: string, endId: string, options?: { sheetId?: string }): AverageFormulaDefinition
  count(startId: string, endId: string, options?: { sheetId?: string }): CountFormulaDefinition
  counta(startId: string, endId: string, options?: { sheetId?: string }): CountAFormulaDefinition
  concat(...values: Array<FormulaDefinition | string>): ConcatenateFormulaDefinition
  iferror(value: FormulaDefinition, fallback: CoercibleValue): IfErrorFormulaDefinition
  vlookup(
    lookup: FormulaDefinition,
    rangeName: string,
    colIndex: number,
    exactMatch?: boolean
  ): VlookupFormulaDefinition

  // --- Generic escape hatch vẫn giữ ---
  call(name: string, ...args: FormulaDefinition[]): CallFormulaDefinition
} as const
```

### Type helper nội bộ

```ts
// Primitive có thể coerce thành literal
type CoercibleValue = FormulaDefinition | string | number | boolean | null;

// Dùng trong implementation — không export
function coerce(value: CoercibleValue): FormulaDefinition {
  if (isFormulaDefinition(value)) return value;
  return { type: 'literal', value };
}
```

`coerce` không export. User không cần biết nó tồn tại.

### So sánh before/after

```ts
// BEFORE — 23 dòng
{
  type: "if",
  condition: {
    type: "binary",
    operator: ">",
    left: {
      type: "round",
      value: {
        type: "sum",
        range: { startId: "score_start", endId: "score_end" },
      },
      digits: 0,
    },
    right: { type: "literal", value: 20 },
  },
  whenTrue: {
    type: "sum",
    values: [
      { type: "ref", id: "bonus" },
      { type: "call", name: "max", args: [
        { type: "ref", id: "score_start" },
        { type: "ref", id: "score_end" },
      ]},
    ],
  },
  whenFalse: { type: "literal", value: 0 },
}

// AFTER — 5 dòng, cùng output
f.if(
  f.gt(f.round(f.sumRange("score_start", "score_end"), 0), 20),
  f.sum(f.ref("bonus"), f.max(f.ref("score_start"), f.ref("score_end"))),
  0
)
```

```ts
// Table accessor — BEFORE
accessor: () => ({
  type: 'binary',
  operator: '*',
  left: { type: 'ref', id: 'hours' },
  right: { type: 'ref', id: 'rate' },
});

// AFTER
accessor: () => f.mul(f.ref('hours'), f.ref('rate'));
```

---

## 2. Named Ranges

### Khái niệm

Named range đặt tên cho một vùng ô trong một sheet. Sau khi compile, Excel UI hiển thị tên trong Name Box, và formula có thể tham chiếu bằng tên thay vì địa chỉ.

Named range phù hợp với **vùng tổng hợp cố định** (grid cells với id đã biết). Không áp dụng cho table data rows vì địa chỉ thay đổi theo data.

### Shape mới trong `WorkbookDefinition`

```ts
interface WorkbookDefinition {
  // ...existing fields...
  namedRanges?: readonly NamedRangeDefinition[];
}

interface NamedRangeDefinition {
  name: string; // Tên Excel-compatible: bắt đầu bằng chữ hoặc _, không có space, không trùng cell address
  sheetId: string; // Sheet chứa vùng này
  startId: string; // cell id của ô bắt đầu
  endId: string; // cell id của ô kết thúc (có thể trùng startId nếu là single cell)
}
```

### Formula type mới

```ts
interface NamedRangeFormulaDefinition {
  type: 'namedRange'
  name: string
}

// Thêm vào FormulaDefinition union
type FormulaDefinition =
  | ...existing variants...
  | NamedRangeFormulaDefinition
```

### Validation rules

- `name` không rỗng, chỉ chứa `[A-Za-z0-9_]`, bắt đầu bằng chữ hoặc `_`.
- `name` không được trùng với cell address hợp lệ (`A1`, `B10`, `AA100`...).
- `name` unique trong workbook (Excel named range scope mặc định là workbook).
- `sheetId` phải tồn tại trong sheet registry.
- `startId` và `endId` phải tồn tại trong workbook formula id registry của sheet `sheetId`.
- Validate sau khi `collectWorkbookFormulaIds` chạy xong — không thể validate sớm hơn vì id registry chưa tồn tại.

### Pipeline thay đổi

```txt
WorkbookDefinition.namedRanges
  → validate (sau collectWorkbookFormulaIds)
  → resolve startId/endId thành CellAddress
  → RenderPlan.namedRanges: ResolvedNamedRange[]
  → ExcelJsWorkbookAdapter ghi ExcelJS workbook.definedNames
  → compileFormula nhận { type: 'namedRange', name } → emit tên trực tiếp
```

### Shape trong RenderPlan

```ts
interface RenderPlan {
  // ...existing fields...
  namedRanges?: ResolvedNamedRange[];
}

interface ResolvedNamedRange {
  name: string;
  sheetName: string; // đã resolve từ sheetId
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}
```

### Formula compile

```ts
case 'namedRange':
  // Validate name tồn tại trong workbook named ranges
  return formula.name  // emit trực tiếp — Excel hiểu tên
```

### ExcelJS adapter

```ts
// Sau khi tạo tất cả worksheets
for (const namedRange of renderPlan.namedRanges ?? []) {
  workbook.definedNames.add(
    `'${namedRange.sheetName}'!$${colName}$${namedRange.startRow}:$${colName2}$${namedRange.endRow}`,
    namedRange.name,
  );
}
```

### Dùng với builder

```ts
defineWorkbook({
  namedRanges: [
    { name: 'SCORE_RANGE', sheetId: 'summary', startId: 'score_start', endId: 'score_end' },
  ],
  sheets: [
    {
      id: 'summary',
      blocks: [
        {
          type: 'grid',
          rows: [
            {
              cells: [
                { id: 'score_start', value: 10 },
                { id: 'score_end', value: 20 },
                {
                  value: f.sum(f.namedRange('SCORE_RANGE')),
                  // compile thành =SUM(SCORE_RANGE)
                },
              ],
            },
          ],
        },
      ],
    },
  ],
});
```

---

## 3. Typed Formula Variants

### Lý do tách khỏi `call`

`{ type: 'call', name: 'max', args: [...] }` không validate argument count, không validate argument types, không có compile-time type safety trên `name`. Typed variants giải quyết tất cả.

### Variants mới thêm vào `FormulaDefinition` union

```ts
// Aggregation với range
interface MaxFormulaDefinition {
  type: 'max';
  values: readonly FormulaDefinition[];
}

interface MinFormulaDefinition {
  type: 'min';
  values: readonly FormulaDefinition[];
}

interface AverageFormulaDefinition {
  type: 'average';
  range: FormulaRangeReference;
}

interface CountFormulaDefinition {
  type: 'count';
  range: FormulaRangeReference;
}

interface CountAFormulaDefinition {
  type: 'counta';
  range: FormulaRangeReference;
}

// String
interface ConcatenateFormulaDefinition {
  type: 'concatenate';
  values: readonly FormulaDefinition[];
}

// Error handling
interface IfErrorFormulaDefinition {
  type: 'iferror';
  value: FormulaDefinition;
  fallback: FormulaDefinition;
}

// Lookup — dùng namedRange thay vì tableId để không phụ thuộc vào layout
interface VlookupFormulaDefinition {
  type: 'vlookup';
  lookup: FormulaDefinition;
  rangeName: string; // tên named range chứa table lookup
  colIndex: number; // 1-based column index trong range
  exactMatch?: boolean; // default true
}
```

### Compile rules

```ts
case 'max':
  return `MAX(${formula.values.map(v => compileFormula(v, context)).join(',')})`

case 'min':
  return `MIN(${formula.values.map(v => compileFormula(v, context)).join(',')})`

case 'average':
  return `AVERAGE(${resolveRange(formula.range, context)})`

case 'count':
  return `COUNT(${resolveRange(formula.range, context)})`

case 'counta':
  return `COUNTA(${resolveRange(formula.range, context)})`

case 'concatenate':
  return `CONCATENATE(${formula.values.map(v => compileFormula(v, context)).join(',')})`

case 'iferror':
  return `IFERROR(${compileFormula(formula.value, context)},${compileFormula(formula.fallback, context)})`

case 'vlookup':
  // Validate rangeName tồn tại trong namedRanges
  return `VLOOKUP(${compileFormula(formula.lookup, context)},${formula.rangeName},${formula.colIndex},${formula.exactMatch === false ? 1 : 0})`
```

### Validation mới trong `validateFormulaDefinition`

Mỗi variant mới cần validation riêng:

- `max`/`min`: `values` là array không rỗng.
- `average`/`count`/`counta`: `range` là `FormulaRangeReference` hợp lệ.
- `concatenate`: `values` là array không rỗng, mỗi phần tử là formula.
- `iferror`: `value` và `fallback` đều là formula.
- `vlookup`: `rangeName` là string không rỗng, `colIndex` là số nguyên dương, `exactMatch` là boolean nếu có.

---

## 4. Formula Result Cache

### Vấn đề

Khi ExcelJS ghi formula cell mà không có `result`, Excel sẽ hiển thị 0 hoặc recalculate toàn bộ workbook khi mở. Với workbook lớn hoặc formula phức tạp, recalculation khi mở rất chậm.

### Thay đổi shape

```ts
// GridCell
interface GridCell extends StyleReference {
  id?: string;
  value?: CellContent;
  formulaResult?: CellValue; // hint cache khi value là FormulaDefinition
  colSpan?: number;
  rowSpan?: number;
  width?: number;
}

// TableSectionCell
interface TableSectionCell<Row = Record<string, unknown>> extends StyleReference {
  // ...existing fields...
  formulaResult?: CellValue;
}

// RenderCell — đã có value làm cache, nhưng cần rõ hơn
interface RenderCell {
  row: number;
  column: number;
  value?: CellValue; // dùng làm cached result khi formula có mặt
  formula?: string;
  formulaResult?: CellValue; // explicit cache — ưu tiên hơn value khi cả hai có
  link?: RenderLink;
  style?: StyleValue;
  inlineStyle?: CellStyleDefinition;
}
```

### Compile rule

Khi `GridCell.value` là `FormulaDefinition` và `GridCell.formulaResult` có mặt:

```ts
builder.addCell(sheetId, {
  row,
  column,
  formula: compiledFormulaString,
  value: cell.formulaResult, // ExcelJS dùng làm cached result
  style: cell.style,
});
```

Khi `formulaResult` không có, `value` trong `RenderCell` để `undefined` — ExcelJS ghi formula không có cache, Excel sẽ recalculate.

### ExcelJS adapter

ExcelJS đã support:

```ts
cell.value = {
  formula: 'SUM(A1:A10)',
  result: 42, // cached result
  date1904: false,
};
```

Không cần thay đổi adapter nếu `RenderCell.value` được set đúng khi formula có cache.

### Warning (deferred)

Không block compile nếu thiếu cache. Warning system sẽ được thêm ở phase riêng — chỉ note trong open decisions.

---

## Implementation Checklist

### Builder (`src/formula.ts`)

- [ ] Tạo `CoercibleValue` type và `coerce` helper nội bộ.
- [ ] Implement tất cả helpers trong `f` object.
- [ ] `f.sumRange` và `f.sum` tách rõ — `sumRange` nhận `startId/endId`, `sum` nhận `...FormulaDefinition[]`.
- [ ] Comparison helpers (`f.gt`, `f.lt`...) coerce `right` tự động.
- [ ] `f.if` coerce `whenTrue` và `whenFalse` tự động.
- [ ] `f.concat` coerce string literal tự động.
- [ ] Export `f` từ `src/index.ts`.
- [ ] Viết tests: mỗi helper trả về đúng object shape, coerce hoạt động đúng, mix builder + object literal compile thành cùng output.

### Named Ranges

- [ ] Thêm `NamedRangeDefinition` interface vào `src/core/types.ts`.
- [ ] Thêm `namedRanges?: readonly NamedRangeDefinition[]` vào `WorkbookDefinition`.
- [ ] Thêm `NamedRangeFormulaDefinition` vào `FormulaDefinition` union.
- [ ] Validate named range names trong `validateWorkbookDefinition` (format, uniqueness).
- [ ] Validate `sheetId` và `startId`/`endId` trong `collectWorkbookFormulaIds` sau khi id registry đầy đủ.
- [ ] Thêm `ResolvedNamedRange` vào `RenderPlan` và `RenderPlanBuilder`.
- [ ] `compileFormula` xử lý `case 'namedRange'`: validate name tồn tại, emit tên trực tiếp.
- [ ] `ExcelJsWorkbookAdapter` ghi `workbook.definedNames` sau khi tất cả sheets được tạo.
- [ ] Thêm `f.namedRange(name)` vào builder.
- [ ] Viết tests: compile formula với named range, validate missing name, adapter ghi đúng defined name, đọc lại file xác nhận.

### Typed Formula Variants

- [ ] Thêm 8 interface mới vào `FormulaDefinition` union trong `src/core/types.ts`.
- [ ] Thêm `case` mới vào `compileFormula` trong `src/compiler/formula-engine.ts`.
- [ ] Thêm `case` mới vào `validateFormulaDefinition` trong `src/core/validation.ts`.
- [ ] Thêm `case` mới vào `collectFormulaSheetIds` trong `src/compiler/index.ts`.
- [ ] Thêm `case` mới vào `assertNever` switch để TypeScript báo lỗi nếu thiếu case.
- [ ] Thêm typed helpers vào `f` object trong `src/formula.ts`.
- [ ] `vlookup` validate `rangeName` tồn tại trong `workbook.namedRanges` tại compile time.
- [ ] Viết tests cho từng variant: compile output đúng, validation đúng, đọc lại file xác nhận.

### Formula Result Cache

- [ ] Thêm `formulaResult?: CellValue` vào `GridCell` trong `src/core/types.ts`.
- [ ] Thêm `formulaResult?: CellValue` vào `TableSectionCell` trong `src/core/types.ts`.
- [ ] Thêm `formulaResult?: CellValue` vào `RenderCell` trong `src/compiler/render-plan.ts`.
- [ ] `compileGridBlock` truyền `formulaResult` vào `builder.addCell` khi có.
- [ ] `compileTableSectionRow` tương tự.
- [ ] `ExcelJsWorkbookAdapter.createFormulaValue` dùng `formulaResult ?? value` làm `result`.
- [ ] Viết tests: cell có formula + formulaResult → Excel đọc lại thấy đúng cached value.

---

## Acceptance

- [ ] Builder `f.*` compile thành output giống hệt object literal tương đương — test bằng `assert.deepEqual`.
- [ ] `f.if(cond, 0, 1)` hoạt động không cần `f.val(0)` — coerce tự động.
- [ ] `f.gt(f.ref('score'), 60)` hoạt động không cần `f.val(60)`.
- [ ] Named range khai báo trong workbook → xuất hiện trong Excel Name Box.
- [ ] Formula `{ type: 'namedRange', name: 'SCORE_RANGE' }` compile thành `=SUM(SCORE_RANGE)` khi dùng trong sum.
- [ ] Missing named range name → lỗi rõ ràng tại compile time.
- [ ] Rename `sheet.name` không ảnh hưởng named range nếu `sheetId` giữ nguyên.
- [ ] `f.max`, `f.min`, `f.average`, `f.count`, `f.iferror`, `f.vlookup` compile ra đúng Excel syntax.
- [ ] `vlookup` với `rangeName` không tồn tại → lỗi rõ ràng.
- [ ] Formula cell có `formulaResult` → Excel đọc lại không recalculate ngay, hiển thị cached value.
- [ ] Tất cả tests hiện tại vẫn pass — không có breaking change.

---

## Open Decisions

- **Named range scope:** Excel support cả workbook-scope và sheet-scope. Phase này chỉ implement workbook-scope (default). Sheet-scope có thể thêm sau bằng `scope?: 'workbook' | 'sheet'` mà không break API.
- **Warning khi thiếu `formulaResult`:** Không implement trong phase này. Để lại cho Diagnostic System phase riêng.
- **`f.sumif`, `f.countif`, `f.hlookup`, `f.index`, `f.match`:** Không include phase này — để lại cho phase formula expansion tiếp theo nếu có nhu cầu thực tế.
- **TypeScript typed `f` với generic workbook:** Hiện tại `f.ref(id)` không validate id tồn tại trong workbook tại compile time. Muốn có thì cần `f` nhận workbook type generic — phức tạp hơn nhiều, để phase sau nếu cần.
- **`f.xref` cross-sheet:** Builder `f.xref(sheetId, id)` không validate `sheetId` tại compile time vì builder không biết workbook context. Validation vẫn xảy ra tại `defineWorkbook` như hiện tại.
