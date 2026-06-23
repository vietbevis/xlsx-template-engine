import { ReportEngineError } from "./errors";
import type { WorkbookDefinition } from "./types";

export function validateWorkbookDefinition(workbook: WorkbookDefinition): void {
  if (!isPlainObject(workbook)) {
    throw new ReportEngineError("Workbook definition must be an object.");
  }

  if (!Array.isArray(workbook.sheets)) {
    throw new ReportEngineError("Workbook definition must include a sheets array.");
  }

  if (workbook.sheets.length === 0) {
    throw new ReportEngineError("Workbook definition must include at least one sheet.");
  }

  const sheetIds = new Set<string>();
  const sheetNames = new Set<string>();

  for (const [index, sheet] of workbook.sheets.entries()) {
    if (!isPlainObject(sheet)) {
      throw new ReportEngineError(`Sheet at index ${index} must be an object.`);
    }

    if (typeof sheet.id !== "string" || sheet.id.trim() === "") {
      throw new ReportEngineError(`Sheet at index ${index} must include a non-empty id.`);
    }

    if (typeof sheet.name !== "string" || sheet.name.trim() === "") {
      throw new ReportEngineError(`Sheet "${sheet.id}" must include a non-empty name.`);
    }

    validateSheetName(sheet.name, sheet.id);

    if (!Array.isArray(sheet.blocks)) {
      throw new ReportEngineError(`Sheet "${sheet.id}" must include a blocks array.`);
    }

    for (const [blockIndex, block] of sheet.blocks.entries()) {
      validateBlock(block, sheet.id, blockIndex);
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

function validateBlock(block: unknown, sheetId: string, blockIndex: number): void {
  if (!isPlainObject(block)) {
    throw new ReportEngineError(`Block ${blockIndex} in sheet "${sheetId}" must be an object.`);
  }

  const blockType = block.type;

  if (typeof blockType !== "string" || blockType.trim() === "") {
    throw new ReportEngineError(`Block ${blockIndex} in sheet "${sheetId}" must include a non-empty type.`);
  }

  switch (blockType) {
    case "title":
      validateCellValue(block.text, `Block ${blockIndex} in sheet "${sheetId}" title text`);
      return;
    case "text":
      validateCellValue(block.text, `Block ${blockIndex} in sheet "${sheetId}" text`);
      return;
    case "spacer":
      if (block.rows !== undefined) {
        const rows = block.rows;

        if (!Number.isInteger(rows) || typeof rows !== "number" || rows < 1) {
          throw new ReportEngineError(`Block ${blockIndex} in sheet "${sheetId}" spacer rows must be a positive integer.`);
        }
      }
      return;
    case "grid":
    case "table":
      return;
    default:
      throw new ReportEngineError(`Unknown block type "${blockType}" in sheet "${sheetId}".`);
  }
}

function validateCellValue(value: unknown, label: string): void {
  if (
    value === null ||
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean" ||
    value instanceof Date
  ) {
    return;
  }

  throw new ReportEngineError(`${label} must be a valid cell value.`);
}

function validateSheetName(sheetName: string, sheetId: string): void {
  if (sheetName.length > 31) {
    throw new ReportEngineError(`Sheet "${sheetId}" name must be 31 characters or fewer.`);
  }

  if (/[:\\/?*[\]]/.test(sheetName)) {
    throw new ReportEngineError(`Sheet "${sheetId}" name contains characters Excel does not allow.`);
  }
}

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}
