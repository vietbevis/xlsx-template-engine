import { ReportEngineError } from "../core/errors";
import type { RenderMergeRange } from "./render-plan";

export interface MergeRange extends RenderMergeRange {
  sheetId: string;
}

export function normalizeMergeRange(
  sheetId: string,
  range: RenderMergeRange,
): MergeRange | null {
  assertPositiveInteger(range.startRow, "merge start row");
  assertPositiveInteger(range.startColumn, "merge start column");
  assertPositiveInteger(range.endRow, "merge end row");
  assertPositiveInteger(range.endColumn, "merge end column");

  if (range.endRow < range.startRow || range.endColumn < range.startColumn) {
    throw new ReportEngineError("Merge range end must be greater than or equal to start.");
  }

  if (range.startRow === range.endRow && range.startColumn === range.endColumn) {
    return null;
  }

  return {
    sheetId,
    startRow: range.startRow,
    startColumn: range.startColumn,
    endRow: range.endRow,
    endColumn: range.endColumn,
  };
}

export function assertMergeDoesNotOverlap(
  candidate: MergeRange,
  existingRanges: MergeRange[],
): void {
  const overlappingRange = existingRanges.find(
    (existing) =>
      existing.sheetId === candidate.sheetId &&
      rangesOverlap(existing, candidate),
  );

  if (overlappingRange) {
    throw new ReportEngineError(
      `Merge range ${formatMergeRange(candidate)} overlaps existing range ${formatMergeRange(overlappingRange)}.`,
    );
  }
}

function rangesOverlap(left: RenderMergeRange, right: RenderMergeRange): boolean {
  return !(
    left.endRow < right.startRow ||
    right.endRow < left.startRow ||
    left.endColumn < right.startColumn ||
    right.endColumn < left.startColumn
  );
}

function formatMergeRange(range: MergeRange): string {
  return `"${range.sheetId}"!R${range.startRow}C${range.startColumn}:R${range.endRow}C${range.endColumn}`;
}

function assertPositiveInteger(value: number, label: string): void {
  if (!Number.isInteger(value) || value < 1) {
    throw new ReportEngineError(`Render plan ${label} must be a positive integer.`);
  }
}
