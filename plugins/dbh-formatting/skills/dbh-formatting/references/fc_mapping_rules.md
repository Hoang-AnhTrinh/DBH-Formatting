# Fc Mapping Rules

Use these rules only when the supplied workbook matches the established DB Head template layout. Inspect a different workbook before applying any cell mapping.

## Source

- Primary source: the axial drawing page or pages explicitly confirmed for the current project.
- Confirm whether both pages use the same Fc table before treating the values as building-wide.
- When native PDF text is incomplete or visually reordered, use rendered image geometry and selective OCR/vision for Fc extraction.

## Excel Layout

- `B26:B29` is merged and should show the confirmed Fc source under the `Fc` label, e.g. `Fc\n(DRAWING-FC-01)`.
- `C26:F29` is the first fixed 4-row block.
- `G26:J29` is the overflow block if more than 4 Fc rows are needed.
- Do not insert rows for this template unless the user explicitly asks.

Each row format is:

```text
<Fc number> <from floor> ～ <to floor>
```

Write only the Fc number, not the `Fc` prefix. Normalize every floor label to `G` in Excel: never write `FL` (for example, `1FL` becomes `1G`). Normalize roof labels `R` and `RSL` to `RG`; never write `R` alone.

## Boundary Convention

When a horizontal boundary is labelled `nG`, treat it as the final floor of the lower Fc block. The block immediately above begins at `(n+1)G`; never assign the boundary floor to both rows or shift it upward. For example, if one band ends at `4G`, write the next band from `5G`.

When the axis drawing uses `SL` labels, an upper boundary `nSL` maps directly to `nG` in Excel. A band that begins at `nSL` starts at `(n+1)G` in Excel. Never subtract one from the upper `SL` value. Symbolically:

| Fc | Source boundary | Excel mapping |
|---:|---|---|
| Fc-A | 基礎～nSL | 基礎～nG |
| Fc-B | nSL～mSL | (n+1)G～mG |
| Fc-C | mSL～RSL | (m+1)G～RG |

The next row starts from the previous row's end floor + 1.

