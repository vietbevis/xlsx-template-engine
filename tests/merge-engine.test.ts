import assert from "node:assert/strict";
import { mkdir } from "node:fs/promises";
import ExcelJS from "exceljs";
import {
  ReportEngineError,
  RenderPlanBuilder,
  normalizeMergeRange,
} from "../src";
import { ExcelJsWorkbookAdapter } from "../src/adapters/exceljs/workbook-adapter";

export async function runMergeEngineTest(): Promise<void> {
  assert.deepEqual(
    normalizeMergeRange("summary", {
      startRow: 1,
      startColumn: 1,
      endRow: 1,
      endColumn: 2,
    }),
    {
      sheetId: "summary",
      startRow: 1,
      startColumn: 1,
      endRow: 1,
      endColumn: 2,
    },
  );

  assert.equal(
    normalizeMergeRange("summary", {
      startRow: 1,
      startColumn: 1,
      endRow: 1,
      endColumn: 1,
    }),
    null,
  );

  assert.throws(
    () =>
      normalizeMergeRange("summary", {
        startRow: 2,
        startColumn: 1,
        endRow: 1,
        endColumn: 1,
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes("end must be greater than or equal to start"),
  );

  const builder = new RenderPlanBuilder();
  builder.addSheet("summary", "Summary");
  builder.addSheet("detail", "Detail");

  builder.addCell("summary", { row: 1, column: 1, value: "Horizontal" });
  builder.addMerge("summary", {
    startRow: 1,
    startColumn: 1,
    endRow: 1,
    endColumn: 3,
  });

  builder.addCell("summary", { row: 2, column: 1, value: "Vertical" });
  builder.addMerge("summary", {
    startRow: 2,
    startColumn: 1,
    endRow: 4,
    endColumn: 1,
  });

  builder.addCell("summary", { row: 2, column: 2, value: "Mixed" });
  builder.addMerge("summary", {
    startRow: 2,
    startColumn: 2,
    endRow: 4,
    endColumn: 3,
  });

  builder.addCell("detail", { row: 1, column: 1, value: "Same range different sheet" });
  builder.addMerge("detail", {
    startRow: 1,
    startColumn: 1,
    endRow: 1,
    endColumn: 3,
  });

  builder.addMerge("summary", {
    startRow: 6,
    startColumn: 1,
    endRow: 6,
    endColumn: 1,
  });

  assert.throws(
    () =>
      builder.addMerge("summary", {
        startRow: 3,
        startColumn: 3,
        endRow: 5,
        endColumn: 4,
      }),
    (error) =>
      error instanceof ReportEngineError &&
      error.message.includes("overlaps existing range"),
  );

  const renderPlan = builder.build();
  assert.equal(renderPlan.sheets[0]?.merges.length, 3);
  assert.equal(renderPlan.sheets[1]?.merges.length, 1);
  assert.deepEqual(renderPlan.sheets[0]?.merges, [
    { startRow: 1, startColumn: 1, endRow: 1, endColumn: 3 },
    { startRow: 2, startColumn: 1, endRow: 4, endColumn: 1 },
    { startRow: 2, startColumn: 2, endRow: 4, endColumn: 3 },
  ]);

  const outputPath = "output/phase-10-merge-engine-demo.xlsx";
  await mkdir("output", { recursive: true });
  await new ExcelJsWorkbookAdapter().writeFile(renderPlan, outputPath);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(outputPath);

  const sheet = renderedWorkbook.getWorksheet("Summary");
  assert.equal(sheet.getCell("A1").value, "Horizontal");
  assert.equal(sheet.getCell("C1").value, "Horizontal");
  assert.equal(sheet.getCell("A4").value, "Vertical");
  assert.equal(sheet.getCell("C4").value, "Mixed");
}
