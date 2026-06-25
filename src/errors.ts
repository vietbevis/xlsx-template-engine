export class ReportEngineError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ReportEngineError';
  }
}

export interface ValidationErrorDetails {
  path?: string;
}

export class ValidationError extends ReportEngineError {
  readonly path?: string;

  constructor(message: string, details: ValidationErrorDetails = {}) {
    super(message);
    this.name = 'ValidationError';
    this.path = details.path;
  }
}

export interface CompileErrorDetails {
  sheetId?: string;
  blockIndex?: number;
  id?: string;
}

abstract class ContextualReportEngineError extends ReportEngineError {
  readonly sheetId?: string;
  readonly blockIndex?: number;
  readonly id?: string;

  protected constructor(name: string, message: string, details: CompileErrorDetails = {}) {
    super(message);
    this.name = name;
    this.sheetId = details.sheetId;
    this.blockIndex = details.blockIndex;
    this.id = details.id;
  }
}

export class CompileError extends ContextualReportEngineError {
  constructor(message: string, details: CompileErrorDetails = {}) {
    super('CompileError', message, details);
  }
}

export class FormulaError extends ContextualReportEngineError {
  constructor(message: string, details: CompileErrorDetails = {}) {
    super('FormulaError', message, details);
  }
}

export class RenderError extends ContextualReportEngineError {
  constructor(message: string, details: CompileErrorDetails = {}) {
    super('RenderError', message, details);
  }
}
