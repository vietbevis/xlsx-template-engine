import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import { createWriteStream } from "node:fs";
import { mkdir } from "node:fs/promises";
import { promisify } from "node:util";
import ExcelJS from "exceljs";
import {
  defineWorkbook,
  renderWorkbook,
} from "../src";

const smallRowCount = 10_000;
const largeRowCount = 100_000;
const execFileAsync = promisify(execFile);

const smallWorkbook = createLargeTableWorkbook(smallRowCount);
const largeWorkbook = createLargeTableWorkbook(largeRowCount);

function createLargeTableWorkbook(rowCount: number) {
  return defineWorkbook({
    styles: {
      number: {
        numFmt: "#,##0",
      },
    },
    sheets: [
      {
        id: "large_table",
        name: "Large Table",
        blocks: [
          {
            type: "table",
            columns: [
              { title: "Index", key: "index", style: "number" },
              { title: "Value", key: "value", style: "number" },
              {
                title: "Double",
                key: "double",
                style: "number",
                accessor: () => ({
                  type: "binary",
                  operator: "*",
                  left: { type: "ref", key: "value" },
                  right: { type: "literal", value: 2 },
                }),
              },
            ],
            data: Array.from({ length: rowCount }, (_item, index) => ({
              index: index + 1,
              value: (index + 1) * 3,
            })),
          },
        ],
      },
    ],
  });
}

interface FormulaCellValue {
  formula?: string;
}

export async function runStreamingRendererTest(): Promise<void> {
  await mkdir("output", { recursive: true });

  const filePath = "output/phase-15-streaming-write-file-demo.xlsx";
  await renderWorkbook(smallWorkbook).writeFile(filePath);
  await assertLargeWorkbook(filePath, smallRowCount);

  const streamPath = "output/phase-15-streaming-write-stream-demo.xlsx";
  await renderWorkbook(smallWorkbook).writeStream(createWriteStream(streamPath));
  await assertLargeWorkbook(streamPath, smallRowCount);

  const largeFilePath = "output/phase-15-streaming-write-file-100k-demo.xlsx";
  await renderWorkbook(largeWorkbook).writeFile(largeFilePath);
  await assertLargeWorkbook(largeFilePath, largeRowCount);
}

async function assertLargeWorkbook(filePath: string, rowCount: number): Promise<void> {
  await execFileAsync("unzip", ["-t", filePath]);

  const renderedWorkbook = new ExcelJS.Workbook();
  await renderedWorkbook.xlsx.readFile(filePath);

  const sheet = renderedWorkbook.getWorksheet("Large Table");
  assert.equal(sheet.rowCount, rowCount + 1);
  assert.equal(sheet.getCell("A1").value, "Index");
  assert.equal(sheet.getCell("A2").value, 1);
  assert.equal(sheet.getCell("B2").value, 3);
  assert.equal((sheet.getCell("C2").value as FormulaCellValue).formula, "(B2*2)");
  assert.equal(sheet.getCell(`A${rowCount + 1}`).value, rowCount);
  assert.equal(sheet.getCell(`B${rowCount + 1}`).value, rowCount * 3);
  assert.equal(
    (sheet.getCell(`C${rowCount + 1}`).value as FormulaCellValue).formula,
    `(B${rowCount + 1}*2)`,
  );
}
