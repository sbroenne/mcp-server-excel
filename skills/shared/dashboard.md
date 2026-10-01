# Dashboards and reports

Adapt the layout to the requested result. Do not add charts or rebuild existing
reports for a read-only task.

1. Prepare source data, using [Excel Tables](table.md) when appropriate.
2. Apply suitable [number formats and widths](range.md), preserving intentional
   layouts. Check for clipped values and `#####`.
3. Add [charts](chart.md) in explicit empty areas, or auto-place a vertical stack.
4. Inspect series, totals, filters, and geometry; fix overlap warnings.
5. Capture the layout when an interactive desktop is available. If not, report
   that visual verification was unavailable rather than claiming it was checked.
6. Save and close only when authorized; keep open if requested.

Before adding a chart, use the [chart-building recipes](chart.md#choose-the-source-before-creating)
to choose source fields, preserve detail behind short labels, group periods,
and make units clear. Reuse suitable existing data instead of forcing a helper.

## Example layout

| Area | Content |
|------|---------|
| A1:D10 | Source or summary data |
| A12:F25 | Main chart |
| G12:L25 | Supporting chart |
| A27:F40 | Optional third chart |
| G27:L40 | Optional fourth chart |

Leave gaps between visuals, use consistent chart sizes and meaningful titles,
and format axes by data type. Use a separate detail sheet when a large dataset
would obscure the summary. Do not hide inconvenient data or change a chart's
scale to imply a result the data does not support.
