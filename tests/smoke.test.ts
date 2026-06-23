import assert from "node:assert/strict";
import {
  compileWorkbookToRenderPlan,
  defineWorkbook,
  isWorkbookDefinition,
  renderWorkbook,
} from "../src";

const workbook = defineWorkbook({
  metadata: {
    title: "Smoke Report",
    author: "xlsx-template-engine",
  },
  sheets: [
    {
      id: "summary",
      name: "Summary",
      blocks: [],
    },
  ],
});

assert.equal(isWorkbookDefinition(workbook), true);

const renderPlan = compileWorkbookToRenderPlan(workbook);
assert.equal(renderPlan.metadata?.title, "Smoke Report");
assert.equal(renderPlan.sheets[0]?.id, "summary");
assert.equal(renderPlan.sheets[0]?.name, "Summary");

const renderer = renderWorkbook(workbook);
assert.equal(typeof renderer.writeFile, "function");
assert.equal(typeof renderer.writeBuffer, "function");
assert.equal(typeof renderer.writeStream, "function");
