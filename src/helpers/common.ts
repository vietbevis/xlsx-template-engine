import { ReportEngineError } from '../errors';
import type { Block } from '../types';

export function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

export function assertPositiveInteger(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new ReportEngineError(`${label} must be a positive integer.`);
  }
}

export function assertNever(value: never): never {
  throw new ReportEngineError(`Unsupported block type "${(value as Block).type}".`);
}
