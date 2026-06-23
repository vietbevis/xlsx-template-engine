import { defineWorkbook, type TypedFormulaDefinition } from "../src";

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
                { key: "label", value: "Grand total" },
                {
                  key: "grand_total",
                  value: {
                    type: "sum",
                    values: [
                      { type: "ref", sheetId: "department", key: "total" },
                      { type: "ref", sheetId: "appendix", key: "manual_adjustment" },
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
                { key: "base_hours", value: 30 },
                { key: "rate", value: 12 },
                {
                  key: "total",
                  value: {
                    type: "binary",
                    operator: "*",
                    left: { type: "ref", key: "base_hours" },
                    right: { type: "ref", key: "rate" },
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
              cells: [{ key: "manual_adjustment", value: 5 }],
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
  key: "total",
};

void validSummaryFormula;

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
                { key: "source", value: 1 },
                { key: "target", value: { type: "ref", key: "source" } },
                // @ts-expect-error local formula keys must exist in the current sheet.
                { key: "broken", value: { type: "ref", key: "missing" } },
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
                  key: "grand_total",
                  // @ts-expect-error formula sheetId must be one of the declared sheet ids.
                  value: {
                    type: "sum",
                    values: [
                      { type: "ref", sheetId: "missing_sheet", key: "total" },
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
              cells: [{ key: "total", value: 360 }],
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
                  key: "grand_total",
                  // @ts-expect-error formula key must exist in the selected sheet.
                  value: {
                    type: "sum",
                    values: [
                      { type: "ref", sheetId: "department", key: "missing_total" },
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
              cells: [{ key: "total", value: 360 }],
            },
          ],
        },
      ],
    },
  ],
});
