import assert from "node:assert/strict";
import { mkdir } from "node:fs/promises";
import ExcelJS from "exceljs";
import {
  compileWorkbookToRenderPlan,
  defineWorkbook,
  interpolateCellValue,
  interpolateVariables,
  renderWorkbook,
  resolvePath,
} from "../src";

export async function runVariableEngineTest(): Promise<void> {
  assert.equal(resolvePath({ student: { name: "Nguyen Van A" } }, "student.name"), "Nguyen Van A");
  assert.equal(resolvePath({ students: ["A", "B", "C"] }, "students.length"), 3);
  assert.equal(resolvePath({ students: ["A", "B", "C"] }, "students.1"), "B");
  assert.equal(resolvePath({ student: {} }, "student.name"), undefined);

  assert.equal(
    interpolateVariables("Report {{title}} - {{students.length}} students", {
      workbook: {
        title: "June",
        students: ["A", "B"],
      },
    }),
    "Report June - 2 students",
  );
  assert.equal(
    interpolateVariables("Missing {{student.name}} value", { workbook: {} }),
    "Missing  value",
  );

  const today = new Date("2026-06-23T00:00:00.000Z");
  assert.equal(interpolateCellValue(42, { workbook: { count: 1 } }), 42);
  assert.equal(interpolateCellValue(true, { workbook: { active: false } }), true);
  assert.equal(interpolateCellValue(today, { workbook: { today: "ignored" } }), today);

  const workbook = defineWorkbook({
    context: {
      report: {
        title: "Workbook title",
      },
      students: [{ name: "A" }, { name: "B" }],
      owner: "Workbook owner",
    },
    styles: {
      body: {},
    },
    sheets: [
      {
        id: "summary",
        name: "Summary",
        context: {
          owner: "Sheet owner",
          sheetLabel: "Overview",
        },
        blocks: [
          { type: "title", text: "{{report.title}} - {{sheetLabel}}" },
          {
            type: "text",
            text: "Students: {{students.length}}, Owner: {{owner}}, Missing: {{missing.path}}",
            context: { owner: "Block owner" },
            style: "body",
          },
          {
            type: "grid",
            rows: [
              {
                cells: [
                  { value: "{{student.name}}" },
                  { value: 12 },
                  { value: true },
                  { value: today },
                ],
              },
            ],
            context: {
              student: {
                name: "Nguyen Van A",
              },
            },
          },
          {
            type: "table",
            columns: [
              { title: "{{labels.name}}", key: "name" },
              { title: "Note", key: "note" },
            ],
            data: [
              { name: "Tran Thi B", note: "{{tableNote}}" },
            ],
            context: {
              labels: {
                name: "Lecturer",
              },
              tableNote: "Rendered from block context",
            },
          },
        ],
      },
    ],
  });

  const renderPlan = compileWorkbookToRenderPlan(workbook, {
    context: {
      report: {
        title: "Runtime title",
      },
    },
  });
  const sheetPlan = renderPlan.sheets[0];

  assert.equal(sheetPlan?.rows[0]?.cells[0]?.value, "Runtime title - Overview");
  assert.equal(
    sheetPlan?.rows[1]?.cells[0]?.value,
    "Students: 2, Owner: Block owner, Missing: ",
  );
  assert.deepEqual(
    sheetPlan?.rows[2]?.cells.map((cell) => cell.value),
    ["Nguyen Van A", 12, true, today],
  );
  assert.deepEqual(
    sheetPlan?.rows[3]?.cells.map((cell) => cell.value),
    ["Lecturer", "Note"],
  );
  assert.deepEqual(
    sheetPlan?.rows[4]?.cells.map((cell) => cell.value),
    ["Tran Thi B", "Rendered from block context"],
  );

  const outputPath = "output/phase-11-variable-engine-demo.xlsx";
  await mkdir("output", { recursive: true });
  await renderWorkbook(workbook, {
    context: {
      report: {
        title: "Runtime title",
      },
    },
  }).writeFile(outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  const sheet = renderedWorkbook.getWorksheet("Summary");
  assert.equal(sheet.getCell("A1").value, "Runtime title - Overview");
  assert.equal(sheet.getCell("A2").value, "Students: 2, Owner: Block owner, Missing: ");
  assert.equal(sheet.getCell("A3").value, "Nguyen Van A");
  assert.equal(sheet.getCell("B3").value, 12);
  assert.equal(sheet.getCell("C3").value, true);
  assert.deepEqual(sheet.getCell("D3").value, today);
  assert.equal(sheet.getCell("A4").value, "Lecturer");
  assert.equal(sheet.getCell("B4").value, "Note");
  assert.equal(sheet.getCell("A5").value, "Tran Thi B");
  assert.equal(sheet.getCell("B5").value, "Rendered from block context");
}
