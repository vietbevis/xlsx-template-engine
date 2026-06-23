# 11. Variable Engine

## Goal

Ho tro template text bang `{{path}}` cho title/text/grid/table cell static values, voi context duoc truyen vao luc render.

## Syntax

Supported examples:

```text
{{today}}
{{student.name}}
{{students.length}}
```

## Context Scope

Context merge theo thu tu uu tien:

- Workbook render context.
- Sheet context neu co.
- Block context neu co.

Phase dau co the chi support workbook-level context de API gon, nhung resolver phai de duong mo rong scope.

## Missing Values

Default policy:

- Missing path render thanh empty string.
- Strict mode co the them sau, khong can trong phase dau.

## Implementation Checklist

- [x] Tao path resolver cho object nested va array length.
- [x] Tao interpolate function thay the nhieu variable trong cung string.
- [x] Chi interpolate string values; number/date/boolean giu nguyen.
- [x] Them tests cho nested path, length, missing path.

## Acceptance

- [x] Nested path resolve dung.
- [x] Array length resolve dung.
- [x] Missing variable khong crash o default mode.
