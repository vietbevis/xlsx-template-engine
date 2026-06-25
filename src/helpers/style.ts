import type { CellStyleDefinition } from '../types';
import { isPlainObject } from './common';

export function cloneStylePart(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map((item) => cloneStylePart(item));
  }

  if (!isPlainObject(value)) {
    return value;
  }

  return Object.fromEntries(Object.entries(value).map(([key, childValue]) => [key, cloneStylePart(childValue)]));
}

export function mergeCellStyles(
  base: CellStyleDefinition | undefined,
  override: CellStyleDefinition | undefined,
): CellStyleDefinition | undefined {
  if (!base && !override) return undefined;
  return mergeStylePart(base, override) ?? {};
}

export function mergeStylePart<T extends Record<string, unknown>>(
  base: T | undefined,
  override: T | undefined,
): T | undefined {
  if (!base && !override) return undefined;

  const merged: Record<string, unknown> = { ...(base ?? {}) };

  for (const [key, value] of Object.entries(override ?? {})) {
    const baseValue = merged[key];
    merged[key] =
      isPlainObject(baseValue) && isPlainObject(value) ? mergeStylePart(baseValue, value) : cloneStylePart(value);
  }

  return merged as T;
}
