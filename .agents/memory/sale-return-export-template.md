---
name: Sale Return export must use the original ERP template
description: Sale Return xlsx download must be produced by filling src/assets/sr-template.xlsx, never by generating a fresh SheetJS workbook.
type: constraint
---

The ERP "Sale Return Accounting" import template contains a hidden "Select Option Sheet"
and x14 dropdown data validations on the data sheet. A workbook generated fresh with
SheetJS has identical headers/values but loses those parts, and the ERP import rejects it.

Rule: export via `buildSRWorkbooksFromTemplate()` (src/lib/sr-template-writer.ts), which
loads `src/assets/sr-template.xlsx` with JSZip and replaces only `<sheetData>` in
`xl/worksheets/sheet1.xml`. Never regenerate the workbook from scratch.
