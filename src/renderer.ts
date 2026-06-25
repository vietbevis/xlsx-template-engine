import type { Writable } from 'stream';
import { compileWorkbookToRenderPlan } from './compile';
import type { WorkbookDefinition } from './types';
import { ExcelJsWorkbookRenderer } from './exceljs-workbook';
import type { RenderContext } from './variable-engine';

export interface WorkbookRenderOptions {
  context?: RenderContext;
}

export interface WorkbookRenderer {
  writeFile(filePath: string): Promise<void>;
  writeBuffer(): Promise<Buffer>;
  writeStream(stream: Writable): Promise<void>;
}

export function renderWorkbook(workbook: WorkbookDefinition, options: WorkbookRenderOptions = {}): WorkbookRenderer {
  const renderPlan = compileWorkbookToRenderPlan(workbook, { context: options.context });
  return new ExcelJsWorkbookRenderer(renderPlan);
}
