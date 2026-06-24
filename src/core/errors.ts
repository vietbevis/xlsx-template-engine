export class ReportEngineError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'ReportEngineError';
  }
}
