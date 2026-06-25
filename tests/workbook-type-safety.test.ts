import type { TypedFormulaDefinition, WorkbookDefinition } from '../src';

const workbook = {
  sheets: [
    {
      id: 'main',
      name: 'Main',
      blocks: [{ type: 'grid', rows: [{ cells: [{ id: 'total', value: 1 }] }] }],
    },
  ],
} as const satisfies WorkbookDefinition;

const localRef: TypedFormulaDefinition<typeof workbook, 'main'> = {
  type: 'ref',
  id: 'total',
};

void localRef;
