import type { CompileContext } from '../compiler/block-dispatcher';
import { ReportEngineError } from '../errors';
import type { CellStyleDefinition, StyleValue } from '../types';
import { isPlainObject } from '../utils/common';

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

/**
 * Giải quyết style cuối cùng cho một cell theo thứ tự ưu tiên:
 *   dynamic (từ styleResolver) > static (khai báo trực tiếp) > inlineStyle (từ border block)
 * và đè lên defaultStyle của workbook.
 *
 * @param staticStyle  Style tĩnh khai báo trong định nghĩa cell/column
 * @param dynamicStyle Style động trả về từ styleResolver (ưu tiên cao hơn staticStyle)
 * @param context      CompileContext chứa styleConfig và style registry
 * @param label        Nhãn để hiển thị trong thông báo lỗi
 * @param inlineStyle  Style inline từ border của block (áp dụng cuối cùng, ưu tiên cao nhất)
 */
export function resolveStyle(
  staticStyle: StyleValue | undefined,
  dynamicStyle: StyleValue | undefined,
  context: CompileContext,
  label: string,
  inlineStyle?: CellStyleDefinition,
): CellStyleDefinition | undefined {
  const effectiveStyle = dynamicStyle !== undefined ? dynamicStyle : staticStyle;

  let baseStyle: CellStyleDefinition | undefined;
  if (typeof effectiveStyle === 'string') {
    const registryStyle = context.styleConfig.styles?.[effectiveStyle];
    if (!registryStyle) {
      throw new ReportEngineError(`${label} returned unknown style "${effectiveStyle}".`);
    }
    baseStyle = registryStyle;
  } else {
    baseStyle = effectiveStyle;
  }

  // Merge theo thứ tự: defaultStyle → baseStyle → inlineStyle
  return mergeCellStyles(mergeCellStyles(context.styleConfig.defaultStyle, baseStyle), inlineStyle);
}
