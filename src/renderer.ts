import type { Writable } from 'stream';
import { compileWorkbook } from './compile';
import type { WorkbookDefinition } from './types';
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
  return compileWorkbook(workbook, { context: options.context });
}
