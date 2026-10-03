# Worksheet Tables versus model tables

Create a worksheet Table when requested or needed, not for every rectangular
dataset. Reuse suitable existing Tables instead of recreating them.

Current Table commands and inputs come from CLI help or MCP tool descriptions.

## Conversion and preservation

Check headers, merged cells, effective boundaries, and sorting risks before
conversion. Creation checks can block deterministic problems, but warnings
remain advisory. A safe-to-create result does not establish that conversion
suits every business or layout requirement. Choose header handling deliberately.

Removing a Table keeps its cell data but can break dependent objects. Shrinking
changes membership; it is not permission to clear excluded cells. Appending
must respect existing column order.

Metadata and data reads answer different questions. Ordinary data reads include
filtered rows; explicitly inspect visible rows when the result should reflect
the active filter.

## Native filtering

Keep filtering within the intended Table. Clearing its filters does not clear
every filter on the worksheet. For plain ranges or criteria-based copied output,
use the [range filtering workflow](range.md#ordinary-range-and-advanced-filtering)
without converting the data to a Table first.

Native date groups, comparisons, value lists, colors, and icons are different
filter choices. Excel can normalize criteria, so compare their meaning and
actual visible rows rather than the request's original spelling.
Unavailable criterion reads are not evidence of an empty filter.

## Styling

The Table owns its visual style; ordinary range formatting is not its style
system. Column number formats and totals calculations are separate decisions.
Preserve an existing template unless a change is requested.

## Model and worksheet results

A worksheet Table is not automatically in Power Pivot. Add it to the model,
or load a query to the model, before relying on model relationships or measures.
After worksheet-source edits, follow the
[model refresh sequence](datamodel.md#refresh-is-not-calculation).

A DAX-backed worksheet Table displays a model query result. It is not a DAX
calculated table inside the model. Use returned query data when no worksheet
object is needed, or a [PivotTable](pivottable.md) for interactive filtering.
