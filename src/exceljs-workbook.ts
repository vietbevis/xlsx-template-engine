import ExcelJS from 'exceljs';
import type { Writable } from 'stream';
import { RenderError, ReportEngineError } from './errors';
import { SheetWriter, type SheetWriterConfig } from './sheet-writer';
import type { WriterSheetView } from './types';

/**
 * Thin wrapper around an ExcelJS Workbook.
 *
 * Provides `createSheetWriter()` to hand off to compile, and output methods
 * (`writeFile`, `writeBuffer`, `writeStream`) to serialize the result.
 *
 * Buffered-only: the entire workbook is built in memory before writing.
 */
export class ExcelJsWorkbook {
  private readonly workbook: ExcelJS.Workbook;

  constructor(author?: string) {
    this.workbook = new ExcelJS.Workbook();
    if (author) this.workbook.creator = author;
  }

  /**
   * Add a worksheet and return a SheetWriter that compiles directly into it.
   */
  createSheetWriter(name: string, config: SheetWriterConfig, views?: WriterSheetView[]): SheetWriter {
    const sheet = this.workbook.addWorksheet(name, { views });
    return new SheetWriter(sheet, config);
  }

  async writeFile(filePath: string): Promise<void> {
    try {
      await this.workbook.xlsx.writeFile(filePath);
    } catch (error) {
      throw normalizeRenderError(error);
    }
  }

  async writeBuffer(): Promise<Buffer> {
    try {
      const buffer = await this.workbook.xlsx.writeBuffer();
      return Buffer.from(buffer);
    } catch (error) {
      throw normalizeRenderError(error);
    }
  }

  async writeStream(stream: Writable): Promise<void> {
    try {
      await this.workbook.xlsx.write(stream);
    } catch (error) {
      throw normalizeRenderError(error);
    }
  }
}

function normalizeRenderError(error: unknown): Error {
  if (error instanceof RenderError) return error;
  if (error instanceof ReportEngineError) return new RenderError(error.message);
  return error instanceof Error ? error : new RenderError(String(error));
}
