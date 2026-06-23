import ExcelJS from "exceljs";
import assert from "node:assert/strict";
import { mkdir } from "node:fs/promises";
import {
  ReportEngineError,
  compileWorkbookToRenderPlan,
  defineWorkbook,
  renderWorkbook,
} from "../src";

interface LecturerRow {
  name: string;
  hours: number;
  rate: number;
}

interface ComplexLecturerRow {
  name: string;
  hours: number;
  rate: number;
  overtimeHours: number;
  projects: number;
  points: number;
}

const workbook = defineWorkbook({
  styles: {
    header: {
      font: { bold: true },
      fill: { foregroundColor: { argb: "FFD9EAF7" } },
      alignment: { horizontal: "center", vertical: "middle" },
    },
    body: {
      alignment: { horizontal: "left" },
    },
    number: {
      numberFormat: "#,##0",
      alignment: { horizontal: "right" },
    },
  },
  sheets: [
    {
      id: "nested",
      name: "Nested Header",
      blocks: [
        {
          type: "table",
          headerStyle: "header",
          bodyStyle: "body",
          columns: [
            { title: "Lecturer", key: "name", width: 24 },
            {
              title: "Teaching",
              children: [
                { title: "Hours", key: "hours", width: 12, style: "number" },
                {
                  title: "Amount",
                  accessor: (row: LecturerRow) => row.hours * row.rate,
                  width: 16,
                  style: "number",
                },
              ],
            },
          ],
          data: [
            { name: "Nguyen Van A", hours: 12, rate: 150000 },
            { name: "Tran Thi B", hours: 8, rate: 175000 },
          ],
        },
        { type: "text", text: "After nested table", style: "body" },
      ],
    },
  ],
});

const complexWorkbook = defineWorkbook({
  styles: {
    header: {
      font: { bold: true },
      fill: { foregroundColor: { argb: "FFE2F0D9" } },
      alignment: { horizontal: "center", vertical: "middle", wrapText: true },
      border: { bottom: { style: "thin" } },
    },
    body: {
      alignment: { horizontal: "left" },
    },
    number: {
      numberFormat: "#,##0",
      alignment: { horizontal: "right" },
    },
  },
  sheets: [
    {
      id: "complex_nested",
      name: "Complex Nested Header",
      blocks: [
        {
          type: "table",
          headerStyle: "header",
          bodyStyle: "body",
          columns: [
            { title: "Lecturer", key: "name", width: 24 },
            {
              title: "Workload",
              children: [
                {
                  title: "Teaching",
                  children: [
                    {
                      title: "Regular",
                      children: [
                        {
                          title: "Hours",
                          key: "hours",
                          width: 12,
                          style: "number",
                        },
                        {
                          title: "Amount",
                          accessor: (row: ComplexLecturerRow) =>
                            row.hours * row.rate,
                          width: 16,
                          style: "number",
                        },
                      ],
                    },
                    {
                      title: "Overtime",
                      key: "overtimeHours",
                      width: 12,
                      style: "number",
                    },
                  ],
                },
                {
                  title: "Research",
                  children: [
                    {
                      title: "Projects",
                      key: "projects",
                      width: 12,
                      style: "number",
                    },
                    {
                      title: "Points",
                      key: "points",
                      width: 12,
                      style: "number",
                    },
                  ],
                },
              ],
            },
            {
              title: "Total",
              accessor: (row: ComplexLecturerRow) =>
                row.hours * row.rate + row.overtimeHours * 200000,
              width: 16,
              style: "number",
            },
          ],
          data: [
            {
              name: "Nguyen Van A",
              hours: 12,
              rate: 150000,
              overtimeHours: 3,
              projects: 2,
              points: 18,
            },
            {
              name: "Tran Thi B",
              hours: 8,
              rate: 175000,
              overtimeHours: 5,
              projects: 1,
              points: 14,
            },
          ],
        },
        { type: "text", text: "After complex nested table", style: "body" },
      ],
    },
  ],
});

export async function runHeaderTreeTest(): Promise<void> {
  const renderPlan = compileWorkbookToRenderPlan(workbook);
  const sheetPlan = renderPlan.sheets[0];

  assert.equal(sheetPlan?.rows.length, 5);
  assert.deepEqual(
    sheetPlan.rows[0]?.cells.map((cell) => cell.value),
    ["Lecturer", "Teaching"],
  );
  assert.deepEqual(
    sheetPlan.rows[1]?.cells.map((cell) => cell.value),
    ["Hours", "Amount"],
  );
  assert.deepEqual(
    sheetPlan.rows[2]?.cells.map((cell) => cell.value),
    ["Nguyen Van A", 12, 1800000],
  );
  assert.deepEqual(
    sheetPlan.rows[3]?.cells.map((cell) => cell.value),
    ["Tran Thi B", 8, 1400000],
  );
  assert.equal(sheetPlan.rows[4]?.index, 5);
  assert.equal(sheetPlan.rows[4]?.cells[0]?.value, "After nested table");
  assert.deepEqual(sheetPlan.merges, [
    { startRow: 1, startColumn: 1, endRow: 2, endColumn: 1 },
    { startRow: 1, startColumn: 2, endRow: 1, endColumn: 3 },
  ]);
  assert.deepEqual(sheetPlan.columnWidths, [
    { column: 1, width: 24 },
    { column: 2, width: 12 },
    { column: 3, width: 16 },
  ]);

  const outputPath = "output/phase-9-header-tree-demo.xlsx";
  await mkdir("output", { recursive: true });
  await renderWorkbook(workbook).writeFile(outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  const sheet = renderedWorkbook.getWorksheet("Nested Header");
  assert.equal(sheet.getCell("A1").value, "Lecturer");
  assert.equal(sheet.getCell("B1").value, "Teaching");
  assert.equal(sheet.getCell("B2").value, "Hours");
  assert.equal(sheet.getCell("C2").value, "Amount");
  assert.equal(sheet.getCell("A3").value, "Nguyen Van A");
  assert.equal(sheet.getCell("B3").value, 12);
  assert.equal(sheet.getCell("C3").value, 1800000);
  assert.equal(sheet.getCell("A5").value, "After nested table");

  const complexRenderPlan = compileWorkbookToRenderPlan(complexWorkbook);
  const complexSheetPlan = complexRenderPlan.sheets[0];

  assert.equal(complexSheetPlan?.rows.length, 7);
  assert.deepEqual(
    complexSheetPlan.rows[0]?.cells.map((cell) => cell.value),
    ["Lecturer", "Workload", "Total"],
  );
  assert.deepEqual(
    complexSheetPlan.rows[1]?.cells.map((cell) => cell.value),
    ["Teaching", "Research"],
  );
  assert.deepEqual(
    complexSheetPlan.rows[2]?.cells.map((cell) => cell.value),
    ["Regular", "Overtime", "Projects", "Points"],
  );
  assert.deepEqual(
    complexSheetPlan.rows[3]?.cells.map((cell) => cell.value),
    ["Hours", "Amount"],
  );
  assert.deepEqual(
    complexSheetPlan.rows[4]?.cells.map((cell) => cell.value),
    ["Nguyen Van A", 12, 1800000, 3, 2, 18, 2400000],
  );
  assert.deepEqual(
    complexSheetPlan.rows[5]?.cells.map((cell) => cell.value),
    ["Tran Thi B", 8, 1400000, 5, 1, 14, 2400000],
  );
  assert.equal(complexSheetPlan.rows[6]?.index, 7);
  assert.equal(
    complexSheetPlan.rows[6]?.cells[0]?.value,
    "After complex nested table",
  );
  assert.deepEqual(complexSheetPlan.merges, [
    { startRow: 1, startColumn: 1, endRow: 4, endColumn: 1 },
    { startRow: 1, startColumn: 2, endRow: 1, endColumn: 6 },
    { startRow: 2, startColumn: 2, endRow: 2, endColumn: 4 },
    { startRow: 3, startColumn: 2, endRow: 3, endColumn: 3 },
    { startRow: 3, startColumn: 4, endRow: 4, endColumn: 4 },
    { startRow: 2, startColumn: 5, endRow: 2, endColumn: 6 },
    { startRow: 3, startColumn: 5, endRow: 4, endColumn: 5 },
    { startRow: 3, startColumn: 6, endRow: 4, endColumn: 6 },
    { startRow: 1, startColumn: 7, endRow: 4, endColumn: 7 },
  ]);
  assert.deepEqual(complexSheetPlan.columnWidths, [
    { column: 1, width: 24 },
    { column: 2, width: 12 },
    { column: 3, width: 16 },
    { column: 4, width: 12 },
    { column: 5, width: 12 },
    { column: 6, width: 12 },
    { column: 7, width: 16 },
  ]);

  const complexOutputPath = "output/phase-9-header-tree-complex-demo.xlsx";
  await renderWorkbook(complexWorkbook).writeFile(complexOutputPath);

  const complexRenderedWorkbook = new ExcelJS.Workbook();
  await complexRenderedWorkbook.xlsx.readFile(complexOutputPath);

  const complexSheet = complexRenderedWorkbook.getWorksheet(
    "Complex Nested Header",
  );
  assert.equal(complexSheet.getCell("A1").value, "Lecturer");
  assert.equal(complexSheet.getCell("B1").value, "Workload");
  assert.equal(complexSheet.getCell("G1").value, "Total");
  assert.equal(complexSheet.getCell("B2").value, "Teaching");
  assert.equal(complexSheet.getCell("E2").value, "Research");
  assert.equal(complexSheet.getCell("B3").value, "Regular");
  assert.equal(complexSheet.getCell("D3").value, "Overtime");
  assert.equal(complexSheet.getCell("B4").value, "Hours");
  assert.equal(complexSheet.getCell("C4").value, "Amount");
  assert.equal(complexSheet.getCell("A5").value, "Nguyen Van A");
  assert.equal(complexSheet.getCell("G5").value, 2400000);
  assert.equal(complexSheet.getCell("A7").value, "After complex nested table");

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: "parent_key",
            name: "Parent Key",
            blocks: [
              {
                type: "table",
                columns: [
                  {
                    title: "Parent",
                    key: "parent",
                    children: [{ title: "Leaf", key: "leaf" }],
                  },
                ],
                data: [],
              },
            ],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes("with children must not include key or accessor"),
  );

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: "leaf_missing_key",
            name: "Leaf Missing Key",
            blocks: [
              {
                type: "table",
                columns: [
                  {
                    title: "Parent",
                    children: [{ title: "Leaf" }],
                  },
                ],
                data: [],
              },
            ],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes("leaf column must include a key or accessor"),
  );
}
