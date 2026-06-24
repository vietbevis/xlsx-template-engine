import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import {
  ReportEngineError,
  compileWorkbookToRenderPlan,
  defineWorkbook,
  renderWorkbook,
} from '../src';

interface LecturerRow {
  name: string;
  hours: number;
  rate: number;
  ignoredExtraField: string;
}

const workbook = defineWorkbook({
  styles: {
    header: {
      font: { bold: true },
      fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9EAF7' } },
      alignment: { horizontal: 'center' },
    },
    body: {
      alignment: { horizontal: 'left' },
    },
    amount: {
      numFmt: '#,##0',
      alignment: { horizontal: 'right' },
    },
  },
  sheets: [
    {
      id: 'table',
      name: 'Table',
      blocks: [
        {
          type: 'table',
          headerStyle: 'header',
          headerRowHeights: [22],
          bodyRowHeight: 18,
          bodyStyle: 'body',
          columns: [
            { title: 'Lecturer', id: 'name', width: 24 },
            { title: 'Hours', id: 'hours', width: 12, style: 'amount' },
            {
              title: 'Amount',
              accessor: (row: LecturerRow) => row.hours * row.rate,
              width: 16,
              style: 'amount',
            },
          ],
          data: [
            { name: 'Nguyen Van A', hours: 12, rate: 150000, ignoredExtraField: 'not rendered' },
            { name: 'Tran Thi B', hours: 8, rate: 175000, ignoredExtraField: 'not rendered' },
          ],
        },
        { type: 'text', text: 'After table', style: 'body' },
      ],
    },
  ],
});

export async function runTableBlockTest(): Promise<void> {
  const renderPlan = compileWorkbookToRenderPlan(workbook);
  const sheetPlan = renderPlan.sheets[0];

  assert.equal(sheetPlan?.rows.length, 4);
  assert.deepEqual(
    sheetPlan.rows[0]?.cells.map((cell) => cell.value),
    ['Lecturer', 'Hours', 'Amount'],
  );
  assert.deepEqual(
    sheetPlan.rows[1]?.cells.map((cell) => cell.value),
    ['Nguyen Van A', 12, 1800000],
  );
  assert.deepEqual(
    sheetPlan.rows[2]?.cells.map((cell) => cell.value),
    ['Tran Thi B', 8, 1400000],
  );
  assert.equal(sheetPlan.rows[3]?.index, 4);
  assert.equal(sheetPlan.rows[3]?.cells[0]?.value, 'After table');
  assert.equal(sheetPlan.rows[1]?.cells.length, 3);
  assert.deepEqual(sheetPlan.columnWidths, [
    { column: 1, width: 24 },
    { column: 2, width: 12 },
    { column: 3, width: 16 },
  ]);
  assert.deepEqual(sheetPlan.rowHeights, [
    { row: 1, height: 22 },
    { row: 2, height: 18 },
    { row: 3, height: 18 },
  ]);

  const outputPath = 'output/phase-8-table-block-demo.xlsx';
  await mkdir('output', { recursive: true });
  await renderWorkbook(workbook).writeFile(outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  const sheet = renderedWorkbook.getWorksheet('Table');
  assert.equal(sheet.getCell('A1').value, 'Lecturer');
  assert.equal(sheet.getCell('B1').value, 'Hours');
  assert.equal(sheet.getCell('C1').value, 'Amount');
  assert.equal(sheet.getCell('A2').value, 'Nguyen Van A');
  assert.equal(sheet.getCell('B2').value, 12);
  assert.equal(sheet.getCell('C2').value, 1800000);
  assert.equal(sheet.getCell('A3').value, 'Tran Thi B');
  assert.equal(sheet.getCell('B3').value, 8);
  assert.equal(sheet.getCell('C3').value, 1400000);
  assert.equal(sheet.getCell('D2').value, null);
  assert.equal(sheet.getCell('A4').value, 'After table');
  assert.equal(sheet.getRow(1).height, 22);
  assert.equal(sheet.getRow(2).height, 18);
  assert.equal(sheet.getRow(3).height, 18);
  assert.equal(sheet.getRow(4).height, undefined);
  assert.equal(sheet.getColumn(1).width, 24);
  assert.equal(sheet.getColumn(2).width, 12);
  assert.equal(sheet.getColumn(3).width, 16);

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: 'missing_columns',
            name: 'Missing Columns',
            blocks: [{ type: 'table', columns: [], data: [] }],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes('columns must be a non-empty array'),
  );

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: 'missing_accessor',
            name: 'Missing Accessor',
            blocks: [
              {
                type: 'table',
                columns: [{ title: 'No id' }],
                data: [],
              },
            ],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes('must include an id or accessor'),
  );

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: 'too_many_header_heights',
            name: 'Too Many Header Heights',
            blocks: [
              {
                type: 'table',
                headerRowHeights: [20, 20],
                columns: [{ title: 'Name', id: 'name' }],
                data: [],
              },
            ],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes('headerRowHeights must not exceed header row count'),
  );

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: 'invalid_body_height',
            name: 'Invalid Body Height',
            blocks: [
              {
                type: 'table',
                bodyRowHeight: 0,
                columns: [{ title: 'Name', id: 'name' }],
                data: [],
              },
            ],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes('bodyRowHeight must be greater than 0'),
  );
}
