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

export class CompileError extends ReportEngineError {
  readonly sheetId?: string;
  readonly blockIndex?: number;
  readonly id?: string;

  constructor(message: string, details: CompileErrorDetails = {}) {
    super(message);
    this.name = 'CompileError';
    this.sheetId = details.sheetId;
    this.blockIndex = details.blockIndex;
    this.id = details.id;
  }
}

export class FormulaError extends ReportEngineError {
  readonly sheetId?: string;
  readonly blockIndex?: number;
  readonly id?: string;

  constructor(message: string, details: CompileErrorDetails = {}) {
    super(message);
    this.name = 'FormulaError';
    this.sheetId = details.sheetId;
    this.blockIndex = details.blockIndex;
    this.id = details.id;
  }
}

export class RenderError extends ReportEngineError {
  readonly sheetId?: string;
  readonly blockIndex?: number;
  readonly id?: string;

  constructor(message: string, details: CompileErrorDetails = {}) {
    super(message);
    this.name = 'RenderError';
    this.sheetId = details.sheetId;
    this.blockIndex = details.blockIndex;
    this.id = details.id;
  }
}
