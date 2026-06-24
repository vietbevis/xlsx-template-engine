import {
  defineWorkbook,
  type FormulaRangeReference,
  type GridCell,
  type RefFormulaDefinition,
  type TableColumn,
  type TableSectionCell,
  type TypedFormulaDefinition,
} from "../src";

const workbook = defineWorkbook({
  sheets: [
    {
      id: "summary",
      name: "Summary",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [
                { id: "label", value: "Grand total" },
                {
                  id: "grand_total",
                  value: {
                    type: "sum",
                    values: [
                      { type: "ref", sheetId: "department", id: "total" },
                      { type: "ref", sheetId: "appendix", id: "manual_adjustment" },
                    ],
                  },
                },
              ],
            },
          ],
        },
      ],
    },
    {
      id: "department",
      name: "Department",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [
                { id: "base_hours", value: 30 },
                { id: "rate", value: 12 },
                {
                  id: "total",
                  value: {
                    type: "binary",
                    operator: "*",
                    left: { type: "ref", id: "base_hours" },
                    right: { type: "ref", id: "rate" },
                  },
                },
              ],
            },
          ],
        },
      ],
    },
    {
      id: "appendix",
      name: "Appendix",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [{ id: "manual_adjustment", value: 5 }],
            },
          ],
        },
      ],
    },
  ],
});

type Workbook = typeof workbook;

const validSummaryFormula: TypedFormulaDefinition<Workbook, "summary"> = {
  type: "ref",
  sheetId: "department",
  id: "total",
};

void validSummaryFormula;

const readonlyWorkbook = defineWorkbook({
  styles: {
    money: { numFmt: "#,##0" },
  },
  sheets: [
    {
      id: "readonly_summary",
      name: "Readonly Summary",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [
                { id: "base", value: 10 },
                { id: "bonus", value: 5 },
                {
                  id: "total",
                  value: {
                    type: "sum",
                    values: [
                      { type: "ref", id: "base" },
                      { type: "ref", id: "bonus" },
                    ],
                  },
                  style: "money",
                },
              ],
            },
          ],
        },
      ],
    },
  ],
} as const);

void readonlyWorkbook;

const deprecatedGridCell: GridCell = {
  // @ts-expect-error grid cells use id, not key.
  key: "legacy",
  value: 1,
};

const deprecatedRefFormula: RefFormulaDefinition = {
  type: "ref",
  // @ts-expect-error ref formulas use id, not key.
  key: "legacy",
};

const deprecatedRangeStart: FormulaRangeReference = {
  // @ts-expect-error formula ranges use startId, not startKey.
  startKey: "start",
  endId: "end",
};

const deprecatedRangeEnd: FormulaRangeReference = {
  startId: "start",
  // @ts-expect-error formula ranges use endId, not endKey.
  endKey: "end",
};

const deprecatedSectionCell: TableSectionCell = {
  // @ts-expect-error section cells use columnId, not columnKey.
  columnKey: "amount",
  value: 1,
};

const deprecatedTableColumn: TableColumn<{ amount: number }> = {
  title: "Amount",
  // @ts-expect-error table columns use id, not key.
  key: "amount",
};

void deprecatedGridCell;
void deprecatedRefFormula;
void deprecatedRangeStart;
void deprecatedRangeEnd;
void deprecatedSectionCell;
void deprecatedTableColumn;

defineWorkbook({
  sheets: [
    {
      id: "typed_summary",
      name: "Typed Summary",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [
                { id: "source", value: 1 },
                { id: "target", value: { type: "ref", id: "source" } },
                // @ts-expect-error local formula keys must exist in the current sheet.
                { id: "broken", value: { type: "ref", id: "missing" } },
              ],
            },
          ],
        },
      ],
    },
  ],
});

defineWorkbook({
  sheets: [
    {
      id: "summary",
      name: "Summary",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [
                {
                  id: "grand_total",
                  // @ts-expect-error formula sheetId must be one of the declared sheet ids.
                  value: {
                    type: "sum",
                    values: [
                      { type: "ref", sheetId: "missing_sheet", id: "total" },
                    ],
                  },
                },
              ],
            },
          ],
        },
      ],
    },
    {
      id: "department",
      name: "Department",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [{ id: "total", value: 360 }],
            },
          ],
        },
      ],
    },
  ],
});

defineWorkbook({
  sheets: [
    {
      id: "summary",
      name: "Summary",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [
                {
                  id: "grand_total",
                  // @ts-expect-error formula id must exist in the selected sheet.
                  value: {
                    type: "sum",
                    values: [
                      { type: "ref", sheetId: "department", id: "missing_total" },
                    ],
                  },
                },
              ],
            },
          ],
        },
      ],
    },
    {
      id: "department",
      name: "Department",
      blocks: [
        {
          type: "grid",
          rows: [
            {
              cells: [{ id: "total", value: 360 }],
            },
          ],
        },
      ],
    },
  ],
});
