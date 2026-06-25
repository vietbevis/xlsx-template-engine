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

export function cloneStyle<T extends CellStyleDefinition>(value: T): T {
  return cloneStylePart(value) as T;
}
