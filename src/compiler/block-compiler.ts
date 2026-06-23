import { ReportEngineError } from "../core/errors";
import type {
  Block,
  SheetDefinition,
  StyleRegistry,
  WorkbookDefinition,
} from "../core/types";
import type { RenderPlanBuilder } from "./render-plan-builder";
import type { LayoutCursor } from "./layout-cursor";

export interface SheetContext {
  workbook: WorkbookDefinition;
  sheet: SheetDefinition;
  styles?: StyleRegistry;
}

export type BlockCompiler<TBlock extends Block = Block> = (
  block: TBlock,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
) => void;

export type BlockCompilerRegistry = {
  [TType in Block["type"]]: BlockCompiler<Extract<Block, { type: TType }>>;
};

export const defaultBlockCompilerRegistry: BlockCompilerRegistry = {
  title(block, context, cursor, builder) {
    builder.addCell(context.sheet.id, {
      row: cursor.row,
      column: cursor.column,
      value: block.text,
      style: block.style,
    });

    if (block.height !== undefined) {
      builder.setRowHeight(context.sheet.id, {
        row: cursor.row,
        height: block.height,
      });
    }

    cursor.advanceRows();
  },
  text(block, context, cursor, builder) {
    builder.addCell(context.sheet.id, {
      row: cursor.row,
      column: cursor.column,
      value: block.text,
      style: block.style,
    });

    if (block.height !== undefined) {
      builder.setRowHeight(context.sheet.id, {
        row: cursor.row,
        height: block.height,
      });
    }

    cursor.advanceRows();
  },
  spacer(block, _context, cursor) {
    cursor.advanceRows(block.rows ?? 1);
  },
  grid(block) {
    throwUnsupportedBlock(block.type);
  },
  table(block) {
    throwUnsupportedBlock(block.type);
  },
};

export function compileBlock(
  block: Block,
  context: SheetContext,
  cursor: LayoutCursor,
  builder: RenderPlanBuilder,
  registry: BlockCompilerRegistry = defaultBlockCompilerRegistry,
): void {
  const compiler = registry[block.type] as BlockCompiler | undefined;

  if (!compiler) {
    throw new ReportEngineError(`Unknown block type "${block.type}" in sheet "${context.sheet.id}".`);
  }

  compiler(block, context, cursor, builder);
}

function throwUnsupportedBlock(blockType: Block["type"]): never {
  throw new ReportEngineError(`Block type "${blockType}" is not supported by the compiler yet.`);
}
