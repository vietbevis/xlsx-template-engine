import ExcelJS from "exceljs";
import assert from "node:assert/strict";
import { mkdir } from "node:fs/promises";
import {
  collectFormulaDependencies,
  compileWorkbookToRenderPlan,
  createSheetRegistry,
  defineWorkbook,
  renderWorkbook,
} from "../src";

interface FormulaCellValue {
  formula?: string;
}

const workbook = defineWorkbook({
  styles: {
    header: {
      font: { bold: true },
    },
  },
  sheets: [
    {
      id: "summary",
      name: "Summary",
      context: {
        label: "Summary Sheet",
      },
      blocks: [
        { type: "title", text: "{{label}}", style: "header" },
        {
          type: "grid",
          rows: [
            {
              cells: [
                { key: "label", value: "Grand total" },
                {
                  key: "grand_total",
                  value: {
                    type: "sum",
                    values: [
                      { type: "ref", sheetId: "department", key: "total" },
                      { type: "ref", sheetId: "appendix", key: "manual_adjustment" },
                    ],
                  },
                },
              ],
            },
          ],
        },
      ],
    },
    {
      id: "department",
      name: "Department",
      context: {
        label: "Department Sheet",
      },
      blocks: [
        { type: "title", text: "{{label}}", style: "header" },
        {
          type: "grid",
          rows: [
            {
              cells: [
                { key: "base_hours", value: 30 },
                { key: "rate", value: 12 },
                {
                  key: "total",
                  value: {
                    type: "binary",
                    operator: "*",
                    left: { type: "ref", key: "base_hours" },
                    right: { type: "ref", key: "rate" },
                  },
                },
              ],
            },
          ],
        },
      ],
    },
    {
      id: "appendix",
      name: "Appendix",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [
                { key: "manual_adjustment", value: 5 },
              ],
            },
          ],
        },
      ],
    },
  ],
});

export async function runMultiSheetTest(): Promise<void> {
  const registry = createSheetRegistry(workbook.sheets);
  assert.deepEqual(Array.from(registry.keys()), ["summary", "department", "appendix"]);
  assert.equal(registry.get("department")?.name, "Department");

  const dependencies = collectFormulaDependencies(workbook);
  assert.deepEqual(dependencies.get("summary"), ["department", "appendix"]);
  assert.deepEqual(dependencies.get("department"), []);

  const renderPlan = compileWorkbookToRenderPlan(workbook);
  assert.deepEqual(
    renderPlan.sheets.map((sheet) => ({ id: sheet.id, name: sheet.name })),
    [
      { id: "summary", name: "Summary" },
      { id: "department", name: "Department" },
      { id: "appendix", name: "Appendix" },
    ],
  );
  assert.equal(renderPlan.sheets[0]?.rows[0]?.cells[0]?.value, "Summary Sheet");
  assert.equal(renderPlan.sheets[1]?.rows[0]?.cells[0]?.value, "Department Sheet");
  assert.equal(
    renderPlan.sheets[0]?.rows[1]?.cells[1]?.formula,
    "SUM('Department'!C2,'Appendix'!A1)",
  );
  assert.equal(renderPlan.sheets[1]?.rows[1]?.cells[2]?.formula, "(A2*B2)");
  assert.equal(renderPlan.sheets[2]?.rows[0]?.cells[0]?.value, 5);

  const outputPath = "output/phase-13-multi-sheet-demo.xlsx";
  await mkdir("output", { recursive: true });
  await renderWorkbook(workbook).writeFile(outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  assert.deepEqual(renderedWorkbook.worksheets.map((sheet) => sheet.name), [
    "Summary",
    "Department",
    "Appendix",
  ]);

  const summarySheet = renderedWorkbook.getWorksheet("Summary");
  const departmentSheet = renderedWorkbook.getWorksheet("Department");
  const appendixSheet = renderedWorkbook.getWorksheet("Appendix");

  assert.equal(summarySheet.getCell("A1").value, "Summary Sheet");
  assert.equal(summarySheet.getCell("A2").value, "Grand total");
  assert.equal(
    (summarySheet.getCell("B2").value as FormulaCellValue).formula,
    "SUM('Department'!C2,'Appendix'!A1)",
  );

  assert.equal(departmentSheet.getCell("A1").value, "Department Sheet");
  assert.equal(departmentSheet.getCell("A2").value, 30);
  assert.equal(departmentSheet.getCell("B2").value, 12);
  assert.equal((departmentSheet.getCell("C2").value as FormulaCellValue).formula, "(A2*B2)");

  assert.equal(appendixSheet.getCell("A1").value, 5);

  const renamedWorkbook = defineWorkbook({
    sheets: [
      {
        id: "summary",
        name: "Summary",
        blocks: [
          {
            type: "grid",
            rows: [
              {
                cells: [
                  {
                    value: { type: "ref", sheetId: "department", key: "total" },
                  },
                ],
              },
            ],
          },
        ],
      },
      {
        id: "department",
        name: "Department's Renamed Sheet",
        blocks: [
          {
            type: "grid",
            rows: [
              {
                cells: [{ key: "total", value: 99 }],
              },
            ],
          },
        ],
      },
    ],
  });
  assert.equal(
    compileWorkbookToRenderPlan(renamedWorkbook).sheets[0]?.rows[0]?.cells[0]?.formula,
    "'Department''s Renamed Sheet'!A1",
  );

  assert.throws(
    () =>
      compileWorkbookToRenderPlan(
        defineWorkbook({
          sheets: [
            {
              id: "summary",
              name: "Summary",
              blocks: [
                {
                  type: "grid",
                  rows: [
                    {
                      cells: [
                        {
                          value: { type: "ref", sheetId: "missing", key: "total" },
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        }),
      ),
    /unknown sheetId "missing"/,
  );
}
