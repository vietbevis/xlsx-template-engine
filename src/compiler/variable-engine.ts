export type RenderContext = Record<string, unknown>;

export interface VariableScope {
  workbook?: RenderContext;
  sheet?: RenderContext;
  block?: RenderContext;
}

const VARIABLE_PATTERN = /\{\{\s*([^{}\s]+)\s*\}\}/g;

export function resolvePath(context: RenderContext | undefined, path: string): unknown {
  if (!context || path.trim() === "") {
    return undefined;
  }

  return path.split(".").reduce<unknown>((current, segment) => {
    if (current === undefined || current === null || segment === "") {
      return undefined;
    }

    if (segment === "length" && (Array.isArray(current) || typeof current === "string")) {
      return current.length;
    }

    if (Array.isArray(current) && /^\d+$/.test(segment)) {
      return current[Number(segment)];
    }

    if (typeof current === "object" && Object.prototype.hasOwnProperty.call(current, segment)) {
      return (current as Record<string, unknown>)[segment];
    }

    return undefined;
  }, context);
}

export function interpolateVariables(value: string, scope: VariableScope): string {
  return value.replace(VARIABLE_PATTERN, (_match, path: string) => {
    const resolvedValue = resolveScopedPath(scope, path);
    return formatInterpolatedValue(resolvedValue);
  });
}

export function interpolateCellValue<TValue>(value: TValue, scope: VariableScope): TValue {
  if (typeof value !== "string") {
    return value;
  }

  return interpolateVariables(value, scope) as TValue;
}

function resolveScopedPath(scope: VariableScope, path: string): unknown {
  const blockValue = resolvePath(scope.block, path);

  if (blockValue !== undefined) {
    return blockValue;
  }

  const sheetValue = resolvePath(scope.sheet, path);

  if (sheetValue !== undefined) {
    return sheetValue;
  }

  return resolvePath(scope.workbook, path);
}

function formatInterpolatedValue(value: unknown): string {
  if (value === undefined || value === null) {
    return "";
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  return String(value);
}
