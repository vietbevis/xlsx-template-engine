import type { Writable } from 'stream';
import { compileWorkbookToRenderPlan } from './compile';
import type { WorkbookDefinition } from './types';
import type { RenderPlan } from './render-plan';
import { ExcelJsWorkbookAdapter } from './exceljs-workbook';
import type { RenderContext } from './variable-engine';

export interface WorkbookRenderOptions {
  context?: RenderContext;
}

export interface WorkbookOutputAdapter {
  writeFile(renderPlan: RenderPlan, filePath: string): Promise<void>;
  writeBuffer(renderPlan: RenderPlan): Promise<Buffer>;
  writeStream(renderPlan: RenderPlan, stream: Writable): Promise<void>;
}

export interface WorkbookRenderer {
  writeFile(filePath: string): Promise<void>;
  writeBuffer(): Promise<Buffer>;
  writeStream(stream: Writable): Promise<void>;
}

export function renderWorkbook(workbook: WorkbookDefinition, options: WorkbookRenderOptions = {}): WorkbookRenderer {
  const renderPlan = compileWorkbookToRenderPlan(workbook, { context: options.context });
  const adapter = new ExcelJsWorkbookAdapter();

  return {
    writeFile(filePath: string): Promise<void> {
      return adapter.writeFile(renderPlan, filePath);
    },
    writeBuffer(): Promise<Buffer> {
      return adapter.writeBuffer(renderPlan);
    },
    writeStream(stream: Writable): Promise<void> {
      return adapter.writeStream(renderPlan, stream);
    },
  };
}
