import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import ExcelJS from 'exceljs';
import { ReportEngineError, defineWorkbook, renderWorkbook } from '../src';
import { ExcelJsWorkbookAdapter } from '../src/adapters/exceljs/workbook-adapter';

const workbook = defineWorkbook({
  defaultStyle: {
    font: {
      name: 'Times New Roman',
      size: 12,
    },
  },
  styles: {
    title: {
      font: {
        name: 'Arial',
        size: 14,
        bold: true,
        color: { argb: 'FF1F4E78' },
      },
      fill: {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFD9EAF7' },
      },
      border: {
        bottom: {
          style: 'thin',
          color: { argb: 'FF4472C4' },
        },
      },
      alignment: {
        horizontal: 'center',
        vertical: 'middle',
        wrapText: true,
      },
      numFmt: '@',
    },
  },
  sheets: [
    {
      id: 'summary',
      name: 'Summary',
      blocks: [{ type: 'title', text: 'Styled title', style: 'title' }],
    },
  ],
});

const directColumnStyleWorkbook = defineWorkbook({
  defaultStyle: {
    font: {
      name: 'Times New Roman',
      size: 12,
    },
  },
  sheets: [
    {
      id: 'direct_column_style',
      name: 'Direct Column Style',
      blocks: [
        {
          type: 'table',
          columns: [
            {
              title: 'Lecturer',
              id: 'name',
              headerStyle: {
                font: { bold: true, family: 4 },
                fill: {
                  type: 'pattern',
                  pattern: 'solid',
                  fgColor: { argb: 'FFE2F0D9' },
                },
              },
            },
            {
              title: 'Amount',
              id: 'amount',
              style: {
                numFmt: '#,##0',
                alignment: { horizontal: 'right', vertical: 'middle' },
                protection: { locked: false },
              },
            },
          ],
          data: [{ name: 'Nguyen Van A', amount: 1200000 }],
        },
      ],
    },
  ],
});

export async function runStyleRegistryTest(): Promise<void> {
  const outputPath = 'output/phase-5-style-registry-demo.xlsx';
  await mkdir('output', { recursive: true });
  await renderWorkbook(workbook).writeFile(outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  const sheet = renderedWorkbook.getWorksheet('Summary');
  const cell = sheet.getCell('A1');

  assert.equal(cell.value, 'Styled title');
  assert.equal(cell.font.bold, true);
  assert.equal(cell.font.name, 'Arial');
  assert.equal(cell.font.size, 14);
  assert.equal(cell.font.color?.argb, 'FF1F4E78');
  assert.equal(cell.fill.type, 'pattern');
  assert.equal(cell.fill.pattern, 'solid');
  assert.equal(cell.fill.fgColor?.argb, 'FFD9EAF7');
  assert.equal(cell.border.bottom?.style, 'thin');
  assert.equal(cell.border.bottom?.color?.argb, 'FF4472C4');
  assert.equal(cell.alignment.horizontal, 'center');
  assert.equal(cell.alignment.vertical, 'middle');
  assert.equal(cell.alignment.wrapText, true);
  assert.equal(cell.numFmt, '@');

  await assert.rejects(
    () =>
      new ExcelJsWorkbookAdapter().writeBuffer({
        sheets: [
          {
            id: 'summary',
            name: 'Summary',
            rows: [
              {
                index: 1,
                cells: [{ row: 1, column: 1, value: 'Missing style', style: 'missing' }],
              },
            ],
            merges: [],
            columnWidths: [],
            rowHeights: [],
          },
        ],
      }),
    (error) => error instanceof ReportEngineError,
  );

  const directOutputPath = 'output/direct-column-style-demo.xlsx';
  await renderWorkbook(directColumnStyleWorkbook).writeFile(directOutputPath);

  const directRenderedWorkbook = new ExcelJS.Workbook();
  await directRenderedWorkbook.xlsx.readFile(directOutputPath);

  const directSheet = directRenderedWorkbook.getWorksheet('Direct Column Style');
  assert.equal(directSheet.getCell('A1').font.bold, true);
  assert.equal(directSheet.getCell('A1').font.name, 'Times New Roman');
  assert.equal(directSheet.getCell('A1').font.size, 12);
  assert.equal(directSheet.getCell('A1').font.family, 4);
  assert.equal(directSheet.getCell('A1').fill.type, 'pattern');
  assert.equal(directSheet.getCell('A1').fill.fgColor?.argb, 'FFE2F0D9');
  assert.equal(directSheet.getCell('B2').numFmt, '#,##0');
  assert.equal(directSheet.getCell('B2').font.name, 'Times New Roman');
  assert.equal(directSheet.getCell('B2').font.size, 12);
  assert.equal(directSheet.getCell('B2').alignment.horizontal, 'right');
  assert.equal(directSheet.getCell('B2').alignment.vertical, 'middle');
  assert.equal(directSheet.getCell('B2').protection.locked, false);
}
