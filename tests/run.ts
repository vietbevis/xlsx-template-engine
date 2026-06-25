import assert from 'node:assert/strict';
import { mkdtemp, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import ExcelJS from 'exceljs';
import {
  FormulaError,
  ValidationError,
  compileWorkbookToRenderPlan,
  renderWorkbook,
  type WorkbookDefinition,
} from '../src';

const execFileAsync = promisify(execFile);

async function main(): Promise<void> {
  testValidationErrorClass();
  testFormulaErrorClass();
  await testSectionRowCanReferenceLaterCellId();
}

function testValidationErrorClass(): void {
  assert.throws(
    () => compileWorkbookToRenderPlan({ sheets: [] } as unknown as WorkbookDefinition),
    (error) => error instanceof ValidationError && error.message.includes('at least one sheet'),
  );
}

function testFormulaErrorClass(): void {
  assert.throws(
    () =>
      compileWorkbookToRenderPlan({
        sheets: [
          {
            id: 'sheet',
            name: 'Sheet',
            blocks: [{ type: 'grid', rows: [{ cells: [{ value: { type: 'ref', id: 'missing' } }] }] }],
          },
        ],
      }),
    (error) => error instanceof FormulaError && error.message.includes('unknown cell id'),
  );
}

async function testSectionRowCanReferenceLaterCellId(): Promise<void> {
  const workbook: WorkbookDefinition = {
    sheets: [
      {
        id: 'sheet',
        name: 'Sheet',
        blocks: [
          {
            type: 'table-groups',
            columns: [
              { id: 'name', title: 'Name' },
              { id: 'amount', title: 'Amount' },
            ],
            groups: [
              {
                data: [{ name: 'A', amount: 10 }],
                footerRows: [
                  {
                    cells: [
                      { columnId: 'amount', value: { type: 'ref', id: 'section_label' } },
                      { id: 'section_label', columnId: 'name', value: 'Total' },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    ],
  };

  const dir = await mkdtemp(join(tmpdir(), 'xlsx-template-engine-'));
  const filePath = join(dir, 'section-row.xlsx');

  try {
    await renderWorkbook(workbook).writeFile(filePath);
    await execFileAsync('unzip', ['-t', filePath]);

    const rendered = new ExcelJS.Workbook();
    await rendered.xlsx.readFile(filePath);
    const sheet = rendered.getWorksheet('Sheet');

    assert.equal(sheet?.getCell('B3').formula, 'A3');
  } finally {
    await rm(dir, { recursive: true, force: true });
  }
}

main().catch((error: unknown) => {
  console.error(error);
  process.exitCode = 1;
});
