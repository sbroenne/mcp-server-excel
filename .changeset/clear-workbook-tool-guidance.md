---
"excelmcp": patch
---

Make accepted chart settings discoverable in CLI help and MCP descriptions, return actual plotted series and values when reading regular charts or PivotCharts, and target secondary axes correctly for titles and number formats. Clarify slicer scope and provide actionable recovery when definition-only Power Query stages prevent batch refresh.

Preserve existing validation rules when type, operator, or error-style inputs are invalid. Report failed Data Model metadata reads and requested measure-format failures instead of inventing empty results or substituting General. Keep model-refresh error details without misdiagnosing unsupported functionality, and correct percentage-label and legacy axis-selector guidance.

Identify the actual measure format during readback instead of misreporting Decimal as Percentage or hiding format-read failures as General.

Advertise only writable chart data-label positions in CLI help and MCP discovery; Excel's read-only Mixed state is not an input option.
