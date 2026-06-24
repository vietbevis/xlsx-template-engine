import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import {
  compileWorkbookToRenderPlan,
  defineWorkbook,
  renderWorkbook,
  type CellContent,
  type TableColumn,
} from '../src';
import { collectFormulaDependencies } from '../src/advanced';

interface LecturerRow extends Record<string, unknown> {
  order: CellContent;
  name: CellContent;
  income: CellContent;
  hours: CellContent;
  amount?: CellContent;
}

type SectionRow =
  | ReturnType<typeof createDepartmentMarker>
  | ReturnType<typeof createTotalRow>
  | ReturnType<typeof createNumberInWordsRow>;
type TableRow = LecturerRow | SectionRow;

const columns: TableColumn<LecturerRow>[] = [
  { title: 'STT', id: 'order', width: 8, bodyStyle: 'number' },
  { title: 'Họ tên Giảng viên', id: 'name', width: 24 },
  { title: 'Thu nhập', id: 'income', width: 14, bodyStyle: 'number' },
  {
    title: 'Thực tế giảng dạy',
    children: [
      { title: 'VN', id: 'hours', width: 10, bodyStyle: 'number' },
      {
        title: 'Thành tiền',
        id: 'amount',
        accessor: (row: LecturerRow) =>
          row.amount ?? {
            type: 'binary',
            operator: '*',
            left: { type: 'ref', id: 'hours' },
            right: { type: 'literal', value: 100000 },
          },
        width: 14,
        bodyStyle: 'number',
      },
    ],
  },
  { title: 'Ký nhận', accessor: () => null, width: 12 },
];

const workbook = defineWorkbook({
  styles: {
    title: {
      font: { bold: true, size: 12 },
      alignment: { horizontal: 'center', vertical: 'middle' },
    },
    header: {
      font: { bold: true },
      alignment: { horizontal: 'center', vertical: 'middle', wrapText: true },
    },
    section: {
      font: { bold: true },
      alignment: { horizontal: 'center', vertical: 'middle' },
    },
    body: {},
    number: {
      numFmt: '#,##0',
      alignment: { horizontal: 'right' },
    },
    numberInWords: {
      font: { italic: true },
      alignment: { horizontal: 'right' },
    },
  },
  sheets: [
    {
      id: 'summary',
      name: 'Overtime',
      blocks: [
        createOvertimeTable({
          subTitle: 'TỔNG HỢP TẤT CẢ CÁC KHOA',
          data: [
            createDepartmentMarker('I', 'KHOA CNTT'),
            {
              order: 1,
              name: 'KHOA CNTT',
              income: ref('cntt', 'income_total'),
              hours: ref('cntt', 'hours_total'),
              amount: ref('cntt', 'amount_total'),
            },
            createDepartmentMarker('II', 'KHOA MẬT MÃ'),
            {
              order: 2,
              name: 'KHOA MẬT MÃ',
              income: ref('mat_ma', 'income_total'),
              hours: ref('mat_ma', 'hours_total'),
              amount: ref('mat_ma', 'amount_total'),
            },
            createTotalRow('allRows', 'Tổng cộng toàn bộ'),
            createNumberInWordsRow(),
          ],
        }),
      ],
    },
    {
      id: 'cntt',
      name: 'KHOA CNTT',
      blocks: [
        createOvertimeTable({
          subTitle: 'KHOA CÔNG NGHỆ THÔNG TIN',
          data: [
            createDepartmentMarker('I', 'KHOA CNTT'),
            { order: 1, name: 'Phạm Văn Hưởng', income: 22541435, hours: 120 },
            { order: 2, name: 'Nguyễn Văn Phác', income: 30424205, hours: 180 },
            createTotalRow('currentRows', 'Tổng cộng'),
            createNumberInWordsRow(),
          ],
        }),
      ],
    },
    {
      id: 'mat_ma',
      name: 'KHOA MẬT MÃ',
      blocks: [
        createOvertimeTable({
          subTitle: 'KHOA MẬT MÃ',
          data: [
            createDepartmentMarker('II', 'KHOA MẬT MÃ'),
            { order: 1, name: 'Bùi Thu Lâm', income: 31535312, hours: 210 },
            createTotalRow('currentRows', 'Tổng cộng'),
            createNumberInWordsRow(),
          ],
        }),
      ],
    },
  ],
});

export async function runSectionedTableTest(): Promise<void> {
  const dependencies = collectFormulaDependencies(workbook);
  assert.deepEqual(dependencies.get('summary'), ['cntt', 'mat_ma']);
  assert.deepEqual(dependencies.get('cntt'), []);
  assert.deepEqual(dependencies.get('mat_ma'), []);

  const renderPlan = compileWorkbookToRenderPlan(workbook);
  assert.deepEqual(
    renderPlan.sheets.map((sheet) => ({ id: sheet.id, name: sheet.name })),
    [
      { id: 'summary', name: 'Overtime' },
      { id: 'cntt', name: 'KHOA CNTT' },
      { id: 'mat_ma', name: 'KHOA MẬT MÃ' },
    ],
  );

  const summarySheet = renderPlan.sheets[0];
  const cnttSheet = renderPlan.sheets[1];
  const matMaSheet = renderPlan.sheets[2];

  assert.equal(getRenderFormula(summarySheet, 6, 3), "'KHOA CNTT'!C8");
  assert.equal(getRenderFormula(summarySheet, 6, 5), "'KHOA CNTT'!E8");
  assert.equal(getRenderFormula(summarySheet, 8, 3), "'KHOA MẬT MÃ'!C7");
  assert.equal(getRenderFormula(summarySheet, 9, 3), 'SUM(C6:C6,C8:C8)');
  assert.equal(getRenderFormula(cnttSheet, 8, 3), 'SUM(C6:C7)');
  assert.equal(getRenderFormula(matMaSheet, 7, 5), 'SUM(E6:E6)');
  assert.equal(getRenderBorderStyle(summarySheet, 5, 1), 'thin');
  assert.equal(getRenderBorderStyle(summarySheet, 9, 6), 'thin');

  const outputPath = 'output/sectioned-table-demo.xlsx';
  await mkdir('output', { recursive: true });
  await renderWorkbook(workbook).writeFile(outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  assert.deepEqual(
    renderedWorkbook.worksheets.map((sheet) => sheet.name),
    ['Overtime', 'KHOA CNTT', 'KHOA MẬT MÃ'],
  );

  const summary = renderedWorkbook.getWorksheet('Overtime');
  const cntt = renderedWorkbook.getWorksheet('KHOA CNTT');
  const matMa = renderedWorkbook.getWorksheet('KHOA MẬT MÃ');

  assert.equal(summary?.getCell('B5').value, 'KHOA CNTT');
  assert.equal(getFormula(summary?.getCell('C6').value), "'KHOA CNTT'!C8");
  assert.equal(getFormula(summary?.getCell('E6').value), "'KHOA CNTT'!E8");
  assert.equal(getFormula(summary?.getCell('C8').value), "'KHOA MẬT MÃ'!C7");
  assert.equal(getFormula(summary?.getCell('C9').value), 'SUM(C6:C6,C8:C8)');
  assert.equal(summary?.getCell('F9').border.bottom?.style, 'thin');

  assert.equal(cntt?.getCell('B5').value, 'KHOA CNTT');
  assert.equal(getFormula(cntt?.getCell('C8').value), 'SUM(C6:C7)');
  assert.equal(getFormula(cntt?.getCell('E8').value), 'SUM(E6:E7)');

  assert.equal(matMa?.getCell('B5').value, 'KHOA MẬT MÃ');
  assert.equal(getFormula(matMa?.getCell('C7').value), 'SUM(C6:C6)');
  assert.equal(getFormula(matMa?.getCell('E7').value), 'SUM(E6:E6)');
}

function createOvertimeTable(options: { subTitle: string; data: Array<TableRow> }) {
  return {
    type: 'table' as const,
    titleRows: [
      {
        value: 'DANH SÁCH GIẢNG VIÊN VƯỢT GIỜ NĂM 2024 - 2025',
        style: 'title',
      },
      { value: options.subTitle, style: 'title' },
    ],
    headerStyle: 'header',
    bodyStyle: 'body',
    border: 'thin' as const,
    columns,
    data: options.data,
  };
}

function createDepartmentMarker(order: string, title: string) {
  return {
    type: 'section' as const,
    resetRows: true,
    style: 'section',
    cells: [
      { column: 1, value: order },
      { column: 2, value: title, colSpan: 'remaining' },
    ],
  };
}

function createTotalRow(scope: 'currentRows' | 'allRows', label: string) {
  return {
    type: 'section' as const,
    style: 'section',
    cells: [
      { column: 1, value: label, colSpan: 2 },
      {
        id: scope === 'currentRows' ? 'income_total' : undefined,
        columnId: 'income',
        style: 'number',
        value: scopedSum('income', scope),
      },
      {
        id: scope === 'currentRows' ? 'hours_total' : undefined,
        columnId: 'hours',
        style: 'number',
        value: scopedSum('hours', scope),
      },
      {
        id: scope === 'currentRows' ? 'amount_total' : undefined,
        columnId: 'amount',
        style: 'number',
        value: scopedSum('amount', scope),
      },
    ],
  };
}

function createNumberInWordsRow() {
  return {
    type: 'section' as const,
    resetRows: true,
    style: 'numberInWords',
    cells: [
      {
        column: 1,
        value: 'Bằng chữ: một trăm sáu mươi tám triệu đồng',
        colSpan: 'remaining' as const,
      },
    ],
  };
}

function ref(sheetId: string, id: string): CellContent {
  return { type: 'ref', sheetId, id };
}

function scopedSum(id: string, scope: 'currentRows' | 'allRows'): CellContent {
  return { type: 'sum', range: { startId: id, endId: id, scope } };
}

function getRenderFormula(
  sheet: ReturnType<typeof compileWorkbookToRenderPlan>['sheets'][number] | undefined,
  rowIndex: number,
  columnIndex: number,
): string | undefined {
  return sheet?.rows
    .find((row) => row.index === rowIndex)
    ?.cells.find((cell) => cell.column === columnIndex)?.formula;
}

function getRenderBorderStyle(
  sheet: ReturnType<typeof compileWorkbookToRenderPlan>['sheets'][number] | undefined,
  rowIndex: number,
  columnIndex: number,
): string | undefined {
  return sheet?.rows
    .find((row) => row.index === rowIndex)
    ?.cells.find((cell) => cell.column === columnIndex)?.inlineStyle?.border?.top?.style;
}

function getFormula(value: ExcelJS.CellValue | undefined): string | undefined {
  if (value && typeof value === 'object' && 'formula' in value) {
    return value.formula;
  }

  return undefined;
}
