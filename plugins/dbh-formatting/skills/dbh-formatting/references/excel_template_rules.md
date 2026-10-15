# Excel Template Rules

## Sheet

- Read `template_schema.json` and use the target sheet name found there. When the workbook has multiple sheets, record the exact chosen sheet name in the mapping config.
- Do not rename a worksheet merely to fit a pipeline convention.

## Checkbox/Form Control Preservation

The template contains Excel controls such as checkboxes. These live in OOXML package parts such as:

- `xl/drawings/*`
- `xl/drawings/*.vml`
- `xl/ctrlProps/*`

Saving the workbook through `openpyxl` or spreadsheet import/export can remove these parts. For output workbooks that need to preserve checkboxes, copy the original `.xlsx` and patch only the necessary XML parts:

- `xl/workbook.xml` for sheet name.
- `xl/worksheets/sheet*.xml` for cell values.

After writing, compare the list/count of control-related package parts between source and output.

## Notes Area

Preserve `その他特記事項` from the template unless the user explicitly asks to overwrite it. Do not add inferred notes to this section by default.

