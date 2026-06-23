import assert from "node:assert/strict";
import { mkdir } from "node:fs/promises";
import ExcelJS from "exceljs";
import {
  ReportEngineError,
  compileWorkbookToRenderPlan,
  defineWorkbook,
  renderWorkbook,
} from "../src";

const workbook = defineWorkbook({
  styles: {
    label: {
      font: { bold: true },
      fill: { foregroundColor: { argb: "FFE2F0D9" } },
      border: { bottom: { style: "thin" } },
    },
    value: {
      alignment: { wrapText: true },
    },
  },
  sheets: [
    {
      id: "form",
      name: "Form",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              height: 24,
              cells: [
                { value: "Department", style: "label", width: 18 },
                { value: "Computer Science", style: "value", colSpan: 2, width: 24 },
              ],
            },
            {
              height: 30,
              cells: [
                { value: "Notes", style: "label", rowSpan: 2 },
                { value: "Grid supports fixed form sections.", style: "value", colSpan: 2 },
              ],
            },
            {
              cells: [{ value: "Continues after row span", style: "value", colSpan: 2 }],
            },
          ],
        },
        { type: "text", text: "After grid", style: "value" },
      ],
    },
  ],
});

export async function runGridBlockTest(): Promise<void> {
  const renderPlan = compileWorkbookToRenderPlan(workbook);
  const sheetPlan = renderPlan.sheets[0];

  assert.equal(sheetPlan?.rows.length, 4);
  assert.equal(sheetPlan.rows[0]?.cells[0]?.value, "Department");
  assert.equal(sheetPlan.rows[0]?.cells[1]?.column, 2);
  assert.equal(sheetPlan.rows[0]?.cells[1]?.value, "Computer Science");
  assert.equal(sheetPlan.rows[1]?.cells[0]?.value, "Notes");
  assert.equal(sheetPlan.rows[2]?.cells[0]?.column, 2);
  assert.equal(sheetPlan.rows[2]?.cells[0]?.value, "Continues after row span");
  assert.equal(sheetPlan.rows[3]?.index, 4);
  assert.equal(sheetPlan.rows[3]?.cells[0]?.value, "After grid");
  assert.deepEqual(sheetPlan.merges, [
    { startRow: 1, startColumn: 2, endRow: 1, endColumn: 3 },
    { startRow: 2, startColumn: 1, endRow: 3, endColumn: 1 },
    { startRow: 2, startColumn: 2, endRow: 2, endColumn: 3 },
    { startRow: 3, startColumn: 2, endRow: 3, endColumn: 3 },
  ]);
  assert.deepEqual(sheetPlan.columnWidths, [
    { column: 1, width: 18 },
    { column: 2, width: 24 },
  ]);
  assert.deepEqual(sheetPlan.rowHeights, [
    { row: 1, height: 24 },
    { row: 2, height: 30 },
  ]);

  const outputPath = "output/phase-7-grid-block-demo.xlsx";
  await mkdir("output", { recursive: true });
  await renderWorkbook(workbook).writeFile(outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  const sheet = renderedWorkbook.getWorksheet("Form");
  assert.equal(sheet.getCell("A1").value, "Department");
  assert.equal(sheet.getCell("B1").value, "Computer Science");
  assert.equal(sheet.getCell("A2").value, "Notes");
  assert.equal(sheet.getCell("B3").value, "Continues after row span");
  assert.equal(sheet.getCell("A4").value, "After grid");
  assert.equal(sheet.getColumn(1).width, 18);
  assert.equal(sheet.getColumn(2).width, 24);
  assert.equal(sheet.getRow(1).height, 24);
  assert.equal(sheet.getRow(2).height, 30);

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: "invalid",
            name: "Invalid",
            blocks: [
              {
                type: "grid",
                rows: [{ cells: [{ value: "Bad", colSpan: 0 }] }],
              },
            ],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes("colSpan must be a positive integer"),
  );

  assert.throws(
    () =>
      compileWorkbookToRenderPlan({
        sheets: [
          {
            id: "overlap",
            name: "Overlap",
            blocks: [
              {
                type: "grid",
                rows: [
                  {
                    cells: [
                      { value: "A" },
                      { value: "B", rowSpan: 2 },
                    ],
                  },
                  {
                    cells: [{ value: "C", colSpan: 2 }],
                  },
                ],
              },
            ],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes("Grid cell merge ranges must not overlap"),
  );
}
