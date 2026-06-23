# 14. Cross Sheet Features

## Goal

Ho tro dependency graph, cross-sheet formula va sheet hyperlink bang `sheetId`.

## Cross-Sheet Reference

```ts
{
  ref: {
    sheetId: "department_cntt",
    cell: "F10"
  }
}
```

Compiler map sang:

```text
='CNTT'!F10
```

Sheet name phai duoc quote/escape dung Excel syntax.

## Dependency Graph

Graph can biet:

- Sheet nao phu thuoc sheet nao.
- Missing sheetId.
- Circular dependency neu dependency anh huong render order.
- Formula/link dependency cho diagnostics.

## Implementation Checklist

- Tao collector de quet formulas/links trong render plan.
- Validate moi `sheetId` ton tai trong sheet registry.
- Tao graph va topological sort neu render order can doi.
- Implement sheet hyperlink compile bang `sheetId`.

## Acceptance

- `sheetId` duoc map sang quoted sheet name.
- Missing `sheetId` fail ro rang.
- Rename `sheet.name` khong lam hong formula/link neu `id` giu nguyen.

