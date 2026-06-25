import { FormulaResolver } from '../formula/formula-resolver';
import { FormulaCompiler } from '../formula/formula-compiler';
import { writeCell, writeMerge } from '../utils/exceljs-writer';
import { assertGridCellDoesNotOverlap, gridOccupancyKey, markGridOccupied, resolveColSpan } from '../utils/grid-layout';
import { resolveStyle } from '../styles/style-resolver';
import type { Block, GridCell } from '../types';
import type { CompileContext } from './block-dispatcher';

interface GridPlacement {
  cell: GridCell;
  row: number;
  column: number;
  rowSpan: number;
  colSpan: number;
}

const COL_ORIGIN = 1;

export function compileGridBlock(
  block: Extract<Block, { type: 'grid' }>,
  context: CompileContext,
  startRow: number,
): number {
  return new GridBlockCompiler(block, context).compile(startRow);
}

class GridBlockCompiler {
  constructor(
    private readonly block: Extract<Block, { type: 'grid' }>,
    private readonly context: CompileContext,
  ) {}

  compile(startRow: number): number {
    const occupied = new Set<string>();
    const placements: GridPlacement[] = [];

    let rowExtent = this.block.rows.length;

    // Pass 1: layout — tính vị trí các cell và đăng ký ID vào registry
    // (Phải chạy trước Pass 2 để công thức cross-reference hoạt động đúng)
    for (const [rowOffset, gridRow] of this.block.rows.entries()) {
      const absRow = startRow + rowOffset;

      if (gridRow.height !== undefined) {
        this.context.worksheet.getRow(absRow).height = gridRow.height;
      }

      let colOffset = 0;

      for (const cell of gridRow.cells) {
        // Bỏ qua các ô đã bị chiếm bởi rowSpan/colSpan của cell phía trên
        while (occupied.has(gridOccupancyKey(rowOffset, colOffset))) colOffset++;

        const colSpan = resolveColSpan(cell.colSpan, colOffset, this.context.sheetColumnCount);
        const rowSpan = cell.rowSpan ?? 1;
        assertGridCellDoesNotOverlap(occupied, rowOffset, colOffset, rowSpan, colSpan);

        const absCol = COL_ORIGIN + colOffset;

        if (cell.id) {
          this.context.registry.register(this.context.sheet.id, this.context.sheet.name, cell.id, absRow, absCol);
        }

        placements.push({ cell, row: absRow, column: absCol, rowSpan, colSpan });
        markGridOccupied(occupied, rowOffset, colOffset, rowSpan, colSpan);

        // Cập nhật rowExtent nếu rowSpan vượt quá số hàng khai báo
        rowExtent = Math.max(rowExtent, rowOffset + rowSpan);
        colOffset += colSpan;
      }
    }

    // Pass 2: render — tất cả ID đã đăng ký, biên dịch và ghi cell
    this.renderCells(placements);

    return startRow + rowExtent;
  }

  private renderCells(placements: GridPlacement[]): void {
    const formulaCtx = FormulaResolver.createGridContext(this.context.registry, this.context.sheet.id);

    for (const p of placements) {
      const compiled = FormulaCompiler.compileCellContent(p.cell.value, formulaCtx);

      const style = resolveStyle(
        p.cell.style,
        p.cell.styleResolver?.(compiled.value),
        this.context,
        `grid cell "${p.cell.id ?? `${p.row}:${p.column}`}"`,
      );

      // Nếu là công thức: truyền formula + kết quả tĩnh; ngược lại truyền giá trị đã biên dịch
      const valueOrFormula = FormulaCompiler.isFormulaDefinition(p.cell.value)
        ? { formula: compiled.formula, result: p.cell.formulaResult }
        : compiled.value;

      writeCell(this.context.worksheet, p.row, p.column, valueOrFormula, style);

      if (p.cell.width !== undefined) {
        this.context.worksheet.getColumn(p.column).width = p.cell.width;
      }

      if (p.rowSpan > 1 || p.colSpan > 1) {
        writeMerge(this.context.worksheet, p.row, p.column, p.row + p.rowSpan - 1, p.column + p.colSpan - 1);
      }
    }
  }
}
