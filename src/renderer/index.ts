import type { Writable } from 'stream';
import { compileWorkbookToRenderPlan } from '../compiler';
import type { WorkbookDefinition } from '../core/types';
import type { RenderPlan } from '../compiler/render-plan';
import { ExcelJsWorkbookAdapter } from '../adapters/exceljs/workbook-adapter';
import type { RenderContext } from '../compiler/variable-engine';

export interface WorkbookRenderOptions {
  adapter?: WorkbookOutputAdapter;
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

export function renderWorkbook(
  workbook: WorkbookDefinition,
  options: WorkbookRenderOptions = {},
): WorkbookRenderer {
  const renderPlan = compileWorkbookToRenderPlan(workbook, { context: options.context });
  const adapter = options.adapter ?? new ExcelJsWorkbookAdapter();

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
