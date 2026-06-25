import type { TypedFormulaDefinition, WorkbookDefinition } from '../src';

const workbook = {
  sheets: [
    {
      id: 'main',
      name: 'Main',
      blocks: [
        { type: 'grid', rows: [{ cells: [{ id: 'total', value: 1 }] }] },
        {
          type: 'table-groups',
          columns: [{ id: 'amount', title: 'Amount' }],
          groups: [
            {
              data: [{ amount: 1 }],
              footerRows: [{ cells: [{ id: 'section_total', columnId: 'amount', value: 'Total' }] }],
            },
          ],
        },
      ],
    },
  ],
} as const satisfies WorkbookDefinition;

const localRef: TypedFormulaDefinition<typeof workbook, 'main'> = {
  type: 'ref',
  id: 'total',
};

void localRef;

const tableColumnRef: TypedFormulaDefinition<typeof workbook, 'main'> = {
  type: 'ref',
  id: 'amount',
};

const tableSectionRef: TypedFormulaDefinition<typeof workbook, 'main'> = {
  type: 'ref',
  id: 'section_total',
};

void tableColumnRef;
void tableSectionRef;
