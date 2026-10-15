# Reusable Excel Mapping

Use `scripts/write_excel_from_mapping.py` when the Excel template must be reused for a new hồ sơ.

## Files

Create two JSON files per project:

1. `data_json`: normalized facts extracted from email/drawings and confirmed by the user.
2. Mapping config: template path, output path, target sheet, cell mappings, clear ranges, and optional Fc layout.

## Confirmed DB Head template mapping rules

When the supplied template uses the established DB Head layout, preserve `その他特記事項` in `B43:B46`. Generated review notes and citations may use `B47:B50`. Do not clear a note cell unless it is being written by the current mapping.

- Map `有効せい係数` to `I32`.
- Map `杭種類` to `H39`.
- Map a confirmed `検討範囲` to `C16` and its `S-...` reference to `C17`. If the scope is unresolved, write red `要確認` in `C16` and clear `C17`.
- Map `接合部横補強筋` to `G33` and its `S-...` reference to `J33`.
- Do not write to or clear `F17`.
- Write every `要確認` value or note in red text. The writer automatically applies red text to any mapped string containing `要確認`.
- If `検討範囲` is unavailable, map red `要確認` to `C16` through `scope_review`; do not keep the template's prior scope.
- Normalize Fc floors to `G`. Never write `FL`; normalize roof labels `R` and `RSL` to `RG`.
- For Fc bands, a boundary `nG` is the final floor of the lower row; begin the next row at `(n+1)G`. If the source axis uses slab labels, map an upper `nSL` directly to `nG`, and map the next band's start to `(n+1)G`. Example: `基礎～2SL` → `基礎～2G`; then `2SL～6SL` → `3G～6G`. Do not subtract one from an `SL` upper boundary. Confirm the boundary geometry on the axis drawing before mapping.

### Steel-diameter rows

For this layout, steel-diameter values belong only in `E22:E25`; keep `D22:D25` blank.

| Cell | Grade | Required value form |
|---|---|---|
| E22 | SD490 | `Dxx～Dxx` |
| E23 | SD390 | `Dxx～Dxx` |
| E24 | SD345 | `Dxx～Dxx` |
| E25 | SD295 | `Dxx～Dxx` |

Write a grade row only when the source confirms it is applied. Never carry forward `以上` or `以下` as the entered value. Establish both endpoints from the current drawing; if an endpoint cannot be evidenced, write red `要確認` and include the source in the review notes rather than inferring it.

### Anchoring and joint transverse reinforcement

For `有効せい係数`, `必要投影定着長さ`, `定着長さ`, `定着長`, and `機械式定着`, collect every relevant value with its condition and `S-...` position. A single unambiguous value may be mapped to `I32`. If more than one applicable value is present, list the entries in the report/review notes; if they conflict, map red `要確認` to `I32` until the user selects the applicable condition.

For `接合部横補強筋`, inspect a `特記なき場合` default before reading individual `仕口部HOOP` rows. When the explicit default is `2-2-D13@100`, map exactly that string to `G33` and its `S-...` source to `J33`. Any different notation, or an absent default, requires red `要確認` in `G33`, with the observed notation and source listed for confirmation. Do not aggregate individual column-list values into a general value.

## Output file name

Use `{管理番号}-{物件名}-検討フォーマット.xlsx` for the completed copy in the source project folder. Never overwrite the user-supplied template. If a file with that exact output name exists, stop and ask before replacing it or selecting another name.

These are layout-specific defaults. Inspect a new template before applying them.

Example data JSON:

```json
{
  "values": {
    "management_no": "PROJECT-001",
    "pipeline_run_date": "2026-01-01",
    "project_name": "REPLACE_WITH_CONFIRMED_PROJECT_NAME",
    "requester": "REPLACE_WITH_CONFIRMED_REQUESTER",
    "email_received_date": "2026-01-01",
    "return_due_date": "2026-01-11",
    "scope": "REPLACE_WITH_CONFIRMED_SCOPE",
    "scope_source": "DRAWING-SCOPE-01",
    "scope_detail": "REPLACE_WITH_CONFIRMED_DETAIL",
    "joint_transverse_reinforcement": "REPLACE_WITH_CONFIRMED_VALUE",
    "joint_transverse_reinforcement_source": "DRAWING-JOINT-01",
    "anchorage_rule": "REPLACE_WITH_CONFIRMED_VALUE"
  },
  "fc_table": {
    "source_label": "DRAWING-FC-01",
    "rows": [
      { "fc": 33, "from": "基礎", "to": "1G" }
    ]
  },
  "citations": [
    { "text": "Project name and design route: DRAWING-META-01 (PDF p.2)" }
  ]
}
```

Example mapping config:

```json
{
  "template": "path/to/template.xlsx",
  "data_json": "path/to/data.json",
  "output": "path/to/result.xlsx",
  "report": "path/to/report.json",
  "sheet_name": "原紙",
  "replace_FL_with_G": true,
  "cells": {
    "D3": { "source": "management_no", "type": "string" },
    "H3": { "source": "pipeline_run_date", "type": "date" },
    "I32": { "source": "effective_depth_coefficient", "type": "string" },
    "G33": { "source": "joint_transverse_reinforcement", "type": "string" },
    "J33": { "source": "joint_transverse_reinforcement_source", "type": "string" },
    "H39": { "source": "pile_type", "type": "string" }
  },
  "clear_cells": ["C17"],
  "clear_ranges": ["C26:J29"],
  "email_dates": {
    "email_pdf": "path/to/2026.01.01_example-email.pdf",
    "received_cell": "D7",
    "return_due_cell": "H7",
    "return_days": 10
  },
  "scope_review": {
    "cell": "C16",
    "source": "scope",
    "source_cell": "C17",
    "citation_source": "scope_source",
    "unresolved_text": "要確認"
  },
  "review_notes": {
    "cells": ["B47", "B48", "B49", "B50"],
    "include_statuses": ["unclear", "unmapped"],
    "font_color": "FFFF0000",
    "template": "・【要確認】{field}：{note}"
  },
  "citation_notes": {
    "cells": ["B47", "B48", "B49"],
    "template": "・参照：{text}"
  },
  "fc_table": {
    "label_cell": "B26",
    "label_template": "Fc\n({source_label})",
    "blocks": [
      { "rows": [26, 27, 28, 29], "columns": ["C", "D", "E", "F"] },
      { "rows": [26, 27, 28, 29], "columns": ["G", "H", "I", "J"] }
    ]
  }
}
```

Run:

```powershell
python <skill-folder>\scripts\write_excel_from_mapping.py --config <mapping-config.json>
```

The script copies the template and patches only worksheet cell XML. It does not save through `openpyxl` or a workbook serialization library. The report includes missing data, missing target cells, Fc overflow, configured citations, XML validity, ZIP validity, control-part counts, and elapsed time.

Use `review_notes` only when the user asks to surface unresolved items in `その他特記事項`. Use `B47:B50` for the confirmed DB Head layout and preserve `B43:B46`. The writer applies red text to review notes and to any other mapping value containing `要確認`.

## Email-file dates and unresolved scope

When approved email filenames are available, add `email_dates` to the mapping config. Use `email_pdf` for one explicitly selected file, or `email_pdfs` for discovered candidates. The writer reads only file names, accepts `YYYY.MM.DD`, `YYYY-MM-DD`, or `YYYYMMDD`, and maps a shared valid date to `D7` and `D7 + 10 days` to `H7`. Multiple candidates are accepted only when every parsed date is identical; different dates, no candidates, or an invalid date name require red `要確認` for both fields. An explicitly user-supplied `email_pdf` takes precedence over folder discovery. It records the derived dates and source file names in the mapping report. It must not open, render, OCR, cache, or derive these dates from the structural drawing or email content.

Add `scope_review` for the normal DB Head layout. If `values.scope` is empty or absent, it writes red `要確認` to `C16`, clears `C17`, and adds an `unclear` review item. If scope is available, it maps the scope to `C16` and `values.scope_source` to `C17`.

When citations are requested, use `data_json.citations` with `citation_notes` to write one source line per configured note cell in `B47:B50`. Cite drawing number and PDF page unless the user asks for drawing number only. Keep field values clean unless the template has a dedicated source-label cell.

For a new template, inspect the sheet name and target cells first. Do not silently reuse a prior project's mapping; create a new data JSON and mapping config when the layout or project facts differ.

