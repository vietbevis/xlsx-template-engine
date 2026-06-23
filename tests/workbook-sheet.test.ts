import assert from "node:assert/strict";
import {
  ReportEngineError,
  compileWorkbookToRenderPlan,
  createSheetRegistry,
  defineWorkbook,
  isWorkbookDefinition,
} from "../src";

assert.throws(
  () => defineWorkbook({ sheets: [] }),
  (error) =>
    error instanceof ReportEngineError &&
    error.message.includes("at least one sheet"),
);

assert.throws(
  () =>
    defineWorkbook({
      sheets: [
        { id: "summary", name: "Summary", blocks: [] },
        { id: "summary", name: "Summary 2", blocks: [] },
      ],
    }),
  /Duplicate sheet id "summary"/,
);

assert.throws(
  () =>
    defineWorkbook({
      sheets: [
        { id: "summary", name: "Summary", blocks: [] },
        { id: "summary_2", name: "summary", blocks: [] },
      ],
    }),
  /Duplicate sheet name "summary"/,
);

assert.throws(
  () =>
    defineWorkbook({
      sheets: [{ id: "summary", name: "Invalid/Name", blocks: [] }],
    }),
  /characters Excel does not allow/,
);

assert.throws(
  () =>
    defineWorkbook({
      sheets: [{ id: "summary", name: "A".repeat(32), blocks: [] }],
    }),
  /31 characters or fewer/,
);

const workbook = defineWorkbook({
  metadata: {
    title: "Workbook Metadata",
    author: "Reporter",
    company: "KMA",
    subject: "Phase 2",
    keywords: ["excel", "report"],
  },
  sheets: [{ id: "summary", name: "Summary", blocks: [] }],
});

const registry = createSheetRegistry(workbook.sheets);
assert.equal(registry.get("summary")?.name, "Summary");

const renderPlan = compileWorkbookToRenderPlan(workbook);
assert.deepEqual(renderPlan.metadata, workbook.metadata);
assert.notEqual(renderPlan.metadata?.keywords, workbook.metadata?.keywords);

assert.equal(isWorkbookDefinition(workbook), true);
assert.equal(isWorkbookDefinition({ sheets: [] }), false);
