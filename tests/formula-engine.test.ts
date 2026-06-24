import assert from "node:assert/strict";
import { mkdir } from "node:fs/promises";
import ExcelJS from "exceljs";
import {
  ReportEngineError,
  compileWorkbookToRenderPlan,
  defineWorkbook,
  renderWorkbook,
  type FormulaDefinition,
} from "../src";
import { compileFormula, createFormulaCompileContext, formatCellAddress } from "../src/advanced";

interface FormulaCellValue {
  formula?: string;
  result?: unknown;
}

const complexFormula: FormulaDefinition = {
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
      {
        type: "call",
        name: "max",
        args: [
          { type: "ref", id: "score_start" },
          { type: "ref", id: "score_end" },
        ],
      },
    ],
  },
  whenFalse: { type: "literal", value: 0 },
};

const workbook = defineWorkbook({
  sheets: [
    {
      id: "formulas",
      name: "Formulas",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [
                { id: "score_start", value: 10 },
                { id: "bonus", value: 5 },
              ],
            },
            {
              cells: [{ id: "score_end", value: 15 }],
            },
            {
              cells: [
                {
                  id: "decision",
                  value: complexFormula,
                },
              ],
            },
          ],
        },
        {
          type: "table",
          columns: [
            { title: "Hours", id: "hours" },
            { title: "Rate", id: "rate" },
            {
              title: "Amount",
              id: "amount",
              accessor: () => ({
                type: "round",
                value: {
                  type: "binary",
                  operator: "*",
                  left: { type: "ref", id: "hours" },
                  right: { type: "ref", id: "rate" },
                },
                digits: 0,
              }),
            },
          ],
          data: [
            { hours: 12, rate: 150000 },
            { hours: 8, rate: 175000 },
          ],
        },
      ],
    },
  ],
});

export async function runFormulaEngineTest(): Promise<void> {
  const formulaContext = createFormulaCompileContext(
    new Map([
      ["score_start", { row: 1, column: 1 }],
      ["score_end", { row: 2, column: 1 }],
      ["bonus", { row: 1, column: 2 }],
    ]),
  );

  assert.equal(formatCellAddress({ row: 3, column: 28 }), "AB3");
  assert.equal(compileFormula({ type: "ref", id: "bonus" }, formulaContext), "B1");
  assert.equal(
    compileFormula(
      {
        type: "sum",
        range: { startId: "score_start", endId: "score_end" },
      },
      formulaContext,
    ),
    "SUM(A1:A2)",
  );
  assert.equal(
    compileFormula(complexFormula, formulaContext),
    "IF((ROUND(SUM(A1:A2),0)>20),SUM(B1,MAX(A1,A2)),0)",
  );

  const renderPlan = compileWorkbookToRenderPlan(workbook);
  const rows = renderPlan.sheets[0]?.rows;

  assert.equal(rows?.[0]?.cells[0]?.value, 10);
  assert.equal(rows?.[0]?.cells[1]?.value, 5);
  assert.equal(rows?.[2]?.cells[0]?.formula, "IF((ROUND(SUM(A1:A2),0)>20),SUM(B1,MAX(A1,A2)),0)");
  assert.equal(rows?.[4]?.cells[2]?.formula, "ROUND((A5*B5),0)");
  assert.equal(rows?.[5]?.cells[2]?.formula, "ROUND((A6*B6),0)");

  const outputPath = "output/phase-12-formula-engine-demo.xlsx";
  await mkdir("output", { recursive: true });
  await renderWorkbook(workbook).writeFile(outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  const sheet = renderedWorkbook.getWorksheet("Formulas");
  assert.equal(sheet.getCell("A1").value, 10);
  assert.equal(sheet.getCell("B1").value, 5);
  assert.equal(
    (sheet.getCell("A3").value as FormulaCellValue).formula,
    "IF((ROUND(SUM(A1:A2),0)>20),SUM(B1,MAX(A1,A2)),0)",
  );
  assert.equal((sheet.getCell("C5").value as FormulaCellValue).formula, "ROUND((A5*B5),0)");
  assert.equal((sheet.getCell("C6").value as FormulaCellValue).formula, "ROUND((A6*B6),0)");
  assert.notEqual(sheet.getCell("A3").value, "IF((ROUND(SUM(A1:A2),0)>20),SUM(B1,MAX(A1,A2)),0)");

  assert.throws(
    () => compileFormula({ type: "ref", id: "missing" }, formulaContext),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes('unknown cell id "missing"'),
  );

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: "deprecated_formula",
            name: "Deprecated Formula",
            blocks: [
              {
                type: "grid",
                rows: [{ cells: [{ value: { type: "ref", key: "legacy" } }] }],
              },
            ],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes('uses deprecated "key"; use id-based fields'),
  );

  assert.throws(
    () =>
      defineWorkbook({
        sheets: [
          {
            id: "invalid_formula",
            name: "Invalid Formula",
            blocks: [
              {
                type: "grid",
                rows: [{ cells: [{ value: { type: "sum" } }] }],
              },
            ],
          },
        ],
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes("sum formula must include a range or values"),
  );
}
