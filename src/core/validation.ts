import { ReportEngineError } from "./errors";
import type { WorkbookDefinition } from "./types";

const FILL_PATTERNS = new Set([
  "none",
  "solid",
  "darkVertical",
  "darkHorizontal",
  "darkGrid",
  "darkTrellis",
  "darkDown",
  "darkUp",
  "lightVertical",
  "lightHorizontal",
  "lightGrid",
  "lightTrellis",
  "lightDown",
  "lightUp",
  "darkGray",
  "mediumGray",
  "lightGray",
  "gray125",
  "gray0625",
]);

const BORDER_LINE_STYLES = new Set([
  "thin",
  "dotted",
  "hair",
  "medium",
  "double",
  "thick",
  "dashDot",
  "dashDotDot",
  "slantDashDot",
  "mediumDashed",
  "mediumDashDotDot",
  "mediumDashDot",
]);

const HORIZONTAL_ALIGNMENTS = new Set([
  "left",
  "center",
  "right",
  "fill",
  "justify",
  "centerContinuous",
  "distributed",
]);

const VERTICAL_ALIGNMENTS = new Set([
  "top",
  "middle",
  "bottom",
  "distributed",
  "justify",
]);

export function validateWorkbookDefinition(workbook: WorkbookDefinition): void {
  if (!isPlainObject(workbook)) {
    throw new ReportEngineError("Workbook definition must be an object.");
  }

  if (!Array.isArray(workbook.sheets)) {
    throw new ReportEngineError(
      "Workbook definition must include a sheets array.",
    );
  }

  if (workbook.sheets.length === 0) {
    throw new ReportEngineError(
      "Workbook definition must include at least one sheet.",
    );
  }

  validateStyleRegistry(workbook.styles);

  const sheetIds = new Set<string>();
  const sheetNames = new Set<string>();

  for (const [index, sheet] of workbook.sheets.entries()) {
    if (!isPlainObject(sheet)) {
      throw new ReportEngineError(`Sheet at index ${index} must be an object.`);
    }

    if (typeof sheet.id !== "string" || sheet.id.trim() === "") {
      throw new ReportEngineError(
        `Sheet at index ${index} must include a non-empty id.`,
      );
    }

    if (typeof sheet.name !== "string" || sheet.name.trim() === "") {
      throw new ReportEngineError(
        `Sheet "${sheet.id}" must include a non-empty name.`,
      );
    }

    validateSheetName(sheet.name, sheet.id);

    if (!Array.isArray(sheet.blocks)) {
      throw new ReportEngineError(
        `Sheet "${sheet.id}" must include a blocks array.`,
      );
    }

    for (const [blockIndex, block] of sheet.blocks.entries()) {
      validateBlock(block, sheet.id, blockIndex, workbook.styles);
    }

    if (sheetIds.has(sheet.id)) {
      throw new ReportEngineError(`Duplicate sheet id "${sheet.id}".`);
    }

    const normalizedSheetName = sheet.name.trim().toLowerCase();

    if (sheetNames.has(normalizedSheetName)) {
      throw new ReportEngineError(`Duplicate sheet name "${sheet.name}".`);
    }

    sheetIds.add(sheet.id);
    sheetNames.add(normalizedSheetName);
  }
}

function validateBlock(
  block: unknown,
  sheetId: string,
  blockIndex: number,
  styles: WorkbookDefinition["styles"],
): void {
  if (!isPlainObject(block)) {
    throw new ReportEngineError(
      `Block ${blockIndex} in sheet "${sheetId}" must be an object.`,
    );
  }

  const blockType = block.type;

  if (typeof blockType !== "string" || blockType.trim() === "") {
    throw new ReportEngineError(
      `Block ${blockIndex} in sheet "${sheetId}" must include a non-empty type.`,
    );
  }

  switch (blockType) {
    case "title":
      validateBlockText(
        block.text,
        `Block ${blockIndex} in sheet "${sheetId}" title text`,
      );
      validateBlockHeight(
        block.height,
        `Block ${blockIndex} in sheet "${sheetId}" title height`,
      );
      validateStyleReference(
        block.style,
        styles,
        `Block ${blockIndex} in sheet "${sheetId}"`,
      );
      return;
    case "text":
      validateBlockText(
        block.text,
        `Block ${blockIndex} in sheet "${sheetId}" text`,
      );
      validateBlockHeight(
        block.height,
        `Block ${blockIndex} in sheet "${sheetId}" text height`,
      );
      validateStyleReference(
        block.style,
        styles,
        `Block ${blockIndex} in sheet "${sheetId}"`,
      );
      return;
    case "spacer":
      if (block.rows !== undefined) {
        const rows = block.rows;

        if (!Number.isInteger(rows) || typeof rows !== "number" || rows < 1) {
          throw new ReportEngineError(
            `Block ${blockIndex} in sheet "${sheetId}" spacer rows must be a positive integer.`,
          );
        }
      }
      return;
    case "grid":
      validateGridBlock(block, sheetId, blockIndex, styles);
      return;
    case "table":
      validateTableBlock(block, sheetId, blockIndex, styles);
      return;
    default:
      throw new ReportEngineError(
        `Unknown block type "${blockType}" in sheet "${sheetId}".`,
      );
  }
}

function validateTableBlock(
  block: Record<string, unknown>,
  sheetId: string,
  blockIndex: number,
  styles: WorkbookDefinition["styles"],
): void {
  const label = `Block ${blockIndex} in sheet "${sheetId}" table`;

  if (!Array.isArray(block.columns) || block.columns.length === 0) {
    throw new ReportEngineError(`${label} columns must be a non-empty array.`);
  }

  if (!Array.isArray(block.data) && !isAsyncIterable(block.data)) {
    throw new ReportEngineError(
      `${label} data must be an array or async iterable.`,
    );
  }

  validateStyleReference(block.headerStyle, styles, `${label} headerStyle`);
  validateStyleReference(block.bodyStyle, styles, `${label} bodyStyle`);

  for (const [columnIndex, column] of block.columns.entries()) {
    validateTableColumn(column, sheetId, blockIndex, [columnIndex], styles);
  }
}

function validateTableColumn(
  column: unknown,
  sheetId: string,
  blockIndex: number,
  columnPath: number[],
  styles: WorkbookDefinition["styles"],
): void {
  const label = `Table column ${columnPath.join(".")} in block ${blockIndex} of sheet "${sheetId}"`;

  if (!isPlainObject(column)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  if (typeof column.title !== "string" || column.title.trim() === "") {
    throw new ReportEngineError(`${label} title must be a non-empty string.`);
  }

  if (
    column.key !== undefined &&
    (typeof column.key !== "string" || column.key.trim() === "")
  ) {
    throw new ReportEngineError(`${label} key must be a non-empty string.`);
  }

  if (column.accessor !== undefined && typeof column.accessor !== "function") {
    throw new ReportEngineError(`${label} accessor must be a function.`);
  }

  const hasChildren = column.children !== undefined;

  if (hasChildren) {
    if (!Array.isArray(column.children) || column.children.length === 0) {
      throw new ReportEngineError(
        `${label} children must be a non-empty array.`,
      );
    }

    if (column.key !== undefined || column.accessor !== undefined) {
      throw new ReportEngineError(
        `${label} with children must not include key or accessor.`,
      );
    }

    for (const [childIndex, childColumn] of column.children.entries()) {
      validateTableColumn(
        childColumn,
        sheetId,
        blockIndex,
        [...columnPath, childIndex],
        styles,
      );
    }
  } else if (column.key === undefined && column.accessor === undefined) {
    throw new ReportEngineError(
      `${label} leaf column must include a key or accessor.`,
    );
  }

  validatePositiveNumber(column.width, `${label} width`);
  validateStyleReference(column.style, styles, label);
}

function validateGridBlock(
  block: Record<string, unknown>,
  sheetId: string,
  blockIndex: number,
  styles: WorkbookDefinition["styles"],
): void {
  if (!Array.isArray(block.rows)) {
    throw new ReportEngineError(
      `Block ${blockIndex} in sheet "${sheetId}" grid rows must be an array.`,
    );
  }

  for (const [rowIndex, row] of block.rows.entries()) {
    if (!isPlainObject(row)) {
      throw new ReportEngineError(
        `Grid row ${rowIndex} in block ${blockIndex} of sheet "${sheetId}" must be an object.`,
      );
    }

    validateBlockHeight(
      row.height,
      `Grid row ${rowIndex} in block ${blockIndex} of sheet "${sheetId}" height`,
    );

    if (!Array.isArray(row.cells)) {
      throw new ReportEngineError(
        `Grid row ${rowIndex} in block ${blockIndex} of sheet "${sheetId}" cells must be an array.`,
      );
    }

    for (const [cellIndex, cell] of row.cells.entries()) {
      validateGridCell(cell, sheetId, blockIndex, rowIndex, cellIndex, styles);
    }
  }
}

function validateGridCell(
  cell: unknown,
  sheetId: string,
  blockIndex: number,
  rowIndex: number,
  cellIndex: number,
  styles: WorkbookDefinition["styles"],
): void {
  const label = `Grid cell ${cellIndex} in row ${rowIndex} of block ${blockIndex} in sheet "${sheetId}"`;

  if (!isPlainObject(cell)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  validateOptionalKey(cell.key, `${label} key`);

  if (cell.value !== undefined) {
    validateCellContent(cell.value, `${label} value`);
  }

  validatePositiveInteger(cell.colSpan, `${label} colSpan`);
  validatePositiveInteger(cell.rowSpan, `${label} rowSpan`);
  validatePositiveNumber(cell.width, `${label} width`);
  validateStyleReference(cell.style, styles, label);
}

function validateStyleRegistry(styles: unknown): void {
  if (styles === undefined) {
    return;
  }

  if (!isPlainObject(styles)) {
    throw new ReportEngineError("Workbook styles must be an object.");
  }

  for (const [styleName, style] of Object.entries(styles)) {
    if (styleName.trim() === "") {
      throw new ReportEngineError("Workbook style names must be non-empty.");
    }

    validateStyleDefinition(style, `Style "${styleName}"`);
  }
}

function validateStyleReference(
  styleName: unknown,
  styles: WorkbookDefinition["styles"],
  label: string,
): void {
  if (styleName === undefined) {
    return;
  }

  if (typeof styleName !== "string" || styleName.trim() === "") {
    throw new ReportEngineError(`${label} style must be a non-empty string.`);
  }

  if (!styles || !Object.prototype.hasOwnProperty.call(styles, styleName)) {
    throw new ReportEngineError(
      `${label} references unknown style "${styleName}".`,
    );
  }
}

function validateStyleDefinition(style: unknown, label: string): void {
  if (!isPlainObject(style)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  if (style.font !== undefined) validateFontStyle(style.font, `${label} font`);
  if (style.fill !== undefined) validateFillStyle(style.fill, `${label} fill`);
  if (style.border !== undefined)
    validateBorderStyle(style.border, `${label} border`);
  if (style.alignment !== undefined)
    validateAlignmentStyle(style.alignment, `${label} alignment`);

  if (
    style.numberFormat !== undefined &&
    typeof style.numberFormat !== "string"
  ) {
    throw new ReportEngineError(`${label} numberFormat must be a string.`);
  }
}

function validateFontStyle(font: unknown, label: string): void {
  if (!isPlainObject(font)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  if (font.name !== undefined && typeof font.name !== "string") {
    throw new ReportEngineError(`${label} name must be a string.`);
  }

  if (
    font.size !== undefined &&
    (typeof font.size !== "number" || font.size <= 0)
  ) {
    throw new ReportEngineError(`${label} size must be greater than 0.`);
  }

  for (const property of ["bold", "italic", "underline"]) {
    if (font[property] !== undefined && typeof font[property] !== "boolean") {
      throw new ReportEngineError(`${label} ${property} must be a boolean.`);
    }
  }

  if (font.color !== undefined) validateColor(font.color, `${label} color`);
}

function validateFillStyle(fill: unknown, label: string): void {
  if (!isPlainObject(fill)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  if (
    fill.pattern !== undefined &&
    (typeof fill.pattern !== "string" || !FILL_PATTERNS.has(fill.pattern))
  ) {
    throw new ReportEngineError(`${label} pattern is not supported.`);
  }

  if (fill.foregroundColor !== undefined)
    validateColor(fill.foregroundColor, `${label} foregroundColor`);
  if (fill.backgroundColor !== undefined)
    validateColor(fill.backgroundColor, `${label} backgroundColor`);
}

function validateBorderStyle(border: unknown, label: string): void {
  if (!isPlainObject(border)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  for (const side of ["top", "right", "bottom", "left"]) {
    if (border[side] !== undefined)
      validateBorderSideStyle(border[side], `${label} ${side}`);
  }
}

function validateBorderSideStyle(side: unknown, label: string): void {
  if (!isPlainObject(side)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  if (typeof side.style !== "string" || !BORDER_LINE_STYLES.has(side.style)) {
    throw new ReportEngineError(`${label} style is not supported.`);
  }

  if (side.color !== undefined) validateColor(side.color, `${label} color`);
}

function validateAlignmentStyle(alignment: unknown, label: string): void {
  if (!isPlainObject(alignment)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  if (
    alignment.horizontal !== undefined &&
    (typeof alignment.horizontal !== "string" ||
      !HORIZONTAL_ALIGNMENTS.has(alignment.horizontal))
  ) {
    throw new ReportEngineError(`${label} horizontal value is not supported.`);
  }

  if (
    alignment.vertical !== undefined &&
    (typeof alignment.vertical !== "string" ||
      !VERTICAL_ALIGNMENTS.has(alignment.vertical))
  ) {
    throw new ReportEngineError(`${label} vertical value is not supported.`);
  }

  if (
    alignment.wrapText !== undefined &&
    typeof alignment.wrapText !== "boolean"
  ) {
    throw new ReportEngineError(`${label} wrapText must be a boolean.`);
  }
}

function validateColor(color: unknown, label: string): void {
  if (!isPlainObject(color)) {
    throw new ReportEngineError(`${label} must be an object.`);
  }

  if (typeof color.argb !== "string" || !/^[0-9A-Fa-f]{8}$/.test(color.argb)) {
    throw new ReportEngineError(
      `${label} argb must be an 8-character hex string.`,
    );
  }
}

function validateBlockText(value: unknown, label: string): void {
  if (typeof value === "string") {
    return;
  }

  throw new ReportEngineError(`${label} must be a string.`);
}

function validateBlockHeight(value: unknown, label: string): void {
  validatePositiveNumber(value, label);
}

function validatePositiveInteger(value: unknown, label: string): void {
  if (value === undefined) {
    return;
  }

  if (!Number.isInteger(value) || typeof value !== "number" || value < 1) {
    throw new ReportEngineError(`${label} must be a positive integer.`);
  }
}

function validatePositiveNumber(value: unknown, label: string): void {
  if (value === undefined) {
    return;
  }

  if (typeof value !== "number" || value <= 0) {
    throw new ReportEngineError(`${label} must be greater than 0.`);
  }
}

function validateCellContent(value: unknown, label: string): void {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    value instanceof Date
  ) {
    return;
  }

  if (isPlainObject(value) && typeof value.type === "string") {
    validateFormulaDefinition(value, label);
    return;
  }

  throw new ReportEngineError(
    `${label} must be a valid cell value or formula definition.`,
  );
}

function validateFormulaDefinition(
  value: Record<string, unknown>,
  label: string,
): void {
  switch (value.type) {
    case "raw":
      if (
        typeof value.expression !== "string" ||
        value.expression.trim() === ""
      ) {
        throw new ReportEngineError(
          `${label} raw formula expression must be a non-empty string.`,
        );
      }

      if (value.expression.trimStart().startsWith("=")) {
        throw new ReportEngineError(
          `${label} formula expression must not start with '='.`,
        );
      }
      return;
    case "literal":
      if (
        value.value !== null &&
        typeof value.value !== "string" &&
        typeof value.value !== "number" &&
        typeof value.value !== "boolean"
      ) {
        throw new ReportEngineError(
          `${label} literal formula value must be a string, number, boolean, or null.`,
        );
      }
      return;
    case "sum":
      if (value.range === undefined && value.values === undefined) {
        throw new ReportEngineError(
          `${label} sum formula must include a range or values.`,
        );
      }

      if (value.range !== undefined) {
        if (!isPlainObject(value.range)) {
          throw new ReportEngineError(
            `${label} sum formula range must be an object.`,
          );
        }

        validateFormulaRangeReference(
          value.range,
          `${label} sum formula range`,
        );
      }

      if (value.values !== undefined) {
        if (!Array.isArray(value.values) || value.values.length === 0) {
          throw new ReportEngineError(
            `${label} sum formula values must be a non-empty array.`,
          );
        }

        for (const [index, child] of value.values.entries()) {
          validateNestedFormulaDefinition(
            child,
            `${label} sum formula value ${index}`,
          );
        }
      }
      return;
    case "round":
      validateNestedFormulaDefinition(
        value.value,
        `${label} round formula value`,
      );

      if (
        !Number.isInteger(value.digits) ||
        typeof value.digits !== "number" ||
        value.digits < 0
      ) {
        throw new ReportEngineError(
          `${label} round formula digits must be a non-negative integer.`,
        );
      }
      return;
    case "if":
      validateNestedFormulaDefinition(
        value.condition,
        `${label} if formula condition`,
      );
      validateNestedFormulaDefinition(
        value.whenTrue,
        `${label} if formula whenTrue`,
      );
      validateNestedFormulaDefinition(
        value.whenFalse,
        `${label} if formula whenFalse`,
      );
      return;
    case "call":
      if (
        typeof value.name !== "string" ||
        !/^[A-Za-z][A-Za-z0-9_.]*$/.test(value.name)
      ) {
        throw new ReportEngineError(
          `${label} call formula name must be a valid Excel function name.`,
        );
      }

      if (!Array.isArray(value.args)) {
        throw new ReportEngineError(
          `${label} call formula args must be an array.`,
        );
      }

      for (const [index, child] of value.args.entries()) {
        validateNestedFormulaDefinition(
          child,
          `${label} call formula arg ${index}`,
        );
      }
      return;
    case "binary":
      if (
        typeof value.operator !== "string" ||
        !["+", "-", "*", "/", ">", ">=", "<", "<=", "=", "<>"].includes(
          value.operator,
        )
      ) {
        throw new ReportEngineError(
          `${label} binary formula operator is not supported.`,
        );
      }

      validateNestedFormulaDefinition(
        value.left,
        `${label} binary formula left`,
      );
      validateNestedFormulaDefinition(
        value.right,
        `${label} binary formula right`,
      );
      return;
    case "range":
      validateFormulaRangeReference(value, `${label} range formula`);
      return;
    case "ref":
      validateRequiredKey(value.key, `${label} ref formula key`);
      return;
    default:
      throw new ReportEngineError(`${label} formula type is not supported.`);
  }
}

function validateNestedFormulaDefinition(value: unknown, label: string): void {
  if (!isPlainObject(value) || typeof value.type !== "string") {
    throw new ReportEngineError(`${label} must be a formula definition.`);
  }

  validateFormulaDefinition(value, label);
}

function validateFormulaRangeReference(
  value: Record<string, unknown>,
  label: string,
): void {
  validateRequiredKey(value.startKey, `${label} startKey`);
  validateRequiredKey(value.endKey, `${label} endKey`);
}

function validateOptionalKey(value: unknown, label: string): void {
  if (value === undefined) {
    return;
  }

  validateRequiredKey(value, label);
}

function validateRequiredKey(value: unknown, label: string): void {
  if (typeof value !== "string" || value.trim() === "") {
    throw new ReportEngineError(`${label} must be a non-empty string.`);
  }
}

function validateSheetName(sheetName: string, sheetId: string): void {
  if (sheetName.length > 31) {
    throw new ReportEngineError(
      `Sheet "${sheetId}" name must be 31 characters or fewer.`,
    );
  }

  if (/[:\\/?*[\]]/.test(sheetName)) {
    throw new ReportEngineError(
      `Sheet "${sheetId}" name contains characters Excel does not allow.`,
    );
  }
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

function isAsyncIterable(value: unknown): value is AsyncIterable<unknown> {
  return (
    typeof value === "object" && value !== null && Symbol.asyncIterator in value
  );
}
