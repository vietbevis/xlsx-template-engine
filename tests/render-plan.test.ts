import assert from "node:assert/strict";
import {
  LayoutCursor,
  ReportEngineError,
  RenderPlanBuilder,
  compileWorkbookToRenderPlan,
  defineWorkbook,
} from "../src";

const workbook = defineWorkbook({
  metadata: { title: "Render Plan" },
  styles: {
    title: { bold: true },
  },
  sheets: [{ id: "summary", name: "Summary", blocks: [] }],
});

const renderPlan = compileWorkbookToRenderPlan(workbook);
assert.equal(renderPlan.metadata?.title, "Render Plan");
assert.deepEqual(renderPlan.styles, workbook.styles);
assert.equal(renderPlan.sheets[0]?.id, "summary");
assert.equal(renderPlan.sheets[0]?.name, "Summary");
assert.deepEqual(renderPlan.sheets[0]?.rows, []);
assert.deepEqual(renderPlan.sheets[0]?.merges, []);

const builder = new RenderPlanBuilder({ title: "Builder" }, { header: {} });
builder.addSheet("summary", "Summary");
builder.addCell("summary", { row: 2, column: 1, value: "A", style: "header" });
builder.addCell("summary", { row: 1, column: 2, value: 10 });
builder.addMerge("summary", {
  startRow: 1,
  startColumn: 1,
  endRow: 1,
  endColumn: 2,
});
builder.setColumnWidth("summary", { column: 1, width: 20 });
builder.setRowHeight("summary", { row: 1, height: 24 });

const builtPlan = builder.build();
assert.equal(builtPlan.sheets[0]?.rows[0]?.index, 1);
assert.equal(builtPlan.sheets[0]?.rows[1]?.index, 2);
assert.equal(builtPlan.sheets[0]?.rows[1]?.cells[0]?.value, "A");
assert.equal(builtPlan.sheets[0]?.merges.length, 1);
assert.equal(builtPlan.sheets[0]?.columnWidths[0]?.width, 20);
assert.equal(builtPlan.sheets[0]?.rowHeights[0]?.height, 24);

const cursor = new LayoutCursor();
cursor.advanceRows(2);
cursor.advanceColumns(3);
assert.equal(cursor.row, 3);
assert.equal(cursor.column, 4);

const blockWorkbook = defineWorkbook({
  sheets: [
    {
      id: "summary",
      name: "Summary",
      blocks: [
        { type: "title", text: "Report Title", style: "title" },
        { type: "spacer", rows: 2 },
        { type: "text", text: "After spacer" },
      ],
    },
  ],
});

const blockPlan = compileWorkbookToRenderPlan(blockWorkbook);
assert.equal(blockPlan.sheets[0]?.rows[0]?.index, 1);
assert.equal(blockPlan.sheets[0]?.rows[0]?.cells[0]?.value, "Report Title");
assert.equal(blockPlan.sheets[0]?.rows[0]?.cells[0]?.style, "title");
assert.equal(blockPlan.sheets[0]?.rows[1]?.index, 4);
assert.equal(blockPlan.sheets[0]?.rows[1]?.cells[0]?.value, "After spacer");

assert.throws(
  () =>
    defineWorkbook({
      sheets: [
        {
          id: "summary",
          name: "Summary",
          blocks: [{ type: "unknown", text: "Nope" } as never],
        },
      ],
    }),
  /Unknown block type "unknown"/,
);

assert.throws(
  () =>
    compileWorkbookToRenderPlan({
      sheets: [
        {
          id: "summary",
          name: "Summary",
          blocks: [{ type: "table" }],
        },
      ],
    }),
  (error) =>
    error instanceof ReportEngineError &&
    error.message.includes('Block type "table" is not supported'),
);
