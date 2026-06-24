import ExcelJS from 'exceljs';
import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
import {
  ReportEngineError,
  compileWorkbookToRenderPlan,
  defineWorkbook,
  renderWorkbook,
} from '../src';

const workbook = defineWorkbook({
  styles: {
    title: {
      font: { bold: true },
      alignment: { horizontal: 'center' },
    },
    body: {
      alignment: { wrapText: true },
    },
  },
  sheets: [
    {
      id: 'summary',
      name: 'Summary',
      blocks: [
        { type: 'title', text: 'Basic Blocks', style: 'title', height: 28 },
        { type: 'spacer', rows: 2 },
        { type: 'text', text: 'Rendered after two spacer rows.', style: 'body', height: 36 },
      ],
    },
  ],
});

export async function runBasicBlocksTest(): Promise<void> {
  const renderPlan = compileWorkbookToRenderPlan(workbook);
  const sheetPlan = renderPlan.sheets[0];

  assert.equal(sheetPlan?.rows.length, 2);
  assert.equal(sheetPlan.rows[0]?.index, 1);
  assert.equal(sheetPlan.rows[0]?.cells[0]?.value, 'Basic Blocks');
  assert.equal(sheetPlan.rows[1]?.index, 4);
  assert.equal(sheetPlan.rows[1]?.cells[0]?.value, 'Rendered after two spacer rows.');
  assert.deepEqual(sheetPlan.rowHeights, [
    { row: 1, height: 28 },
    { row: 4, height: 36 },
  ]);

  const outputPath = 'output/phase-6-basic-blocks-demo.xlsx';
  await mkdir('output', { recursive: true });
  await renderWorkbook(workbook).writeFile(outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  const sheet = renderedWorkbook.getWorksheet('Summary');
  assert.equal(sheet.getCell('A1').value, 'Basic Blocks');
  assert.equal(sheet.getRow(1).height, 28);
  assert.equal(sheet.getCell('A2').value, null);
  assert.equal(sheet.getCell('A3').value, null);
  assert.equal(sheet.getCell('A4').value, 'Rendered after two spacer rows.');
  assert.equal(sheet.getRow(4).height, 36);

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: 'summary',
            name: 'Summary',
            blocks: [{ type: 'title', text: 'Invalid', height: 0 }],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes('sheet "summary" > title block 1 height must be greater than 0'),
  );
}
