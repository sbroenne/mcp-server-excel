---
title: Excel Automation Reference for AI Agents
description: The expert reference corpus that ships inside the Excel MCP Server agent skills - COM behaviour, Power Query and DAX gotchas, PivotTable limits, anti-patterns, and sequencing rules.
keywords: "Excel automation reference, Excel COM gotchas, Power Query reference, Excel MCP agent guidance, Excel automation anti-patterns"
---

# Excel Automation Reference

This is the reference corpus that ships inside the
[Excel MCP Server agent skills](../skills.md) and as MCP prompts. It is written as
instruction for an AI agent — terse, imperative, and specific about what Excel
actually does rather than what its documentation implies.

It is published here because the same material is useful to anyone automating
Excel, whether through this project or not.

For task walkthroughs, start with the [guides](../guides/index.md). For the
complete operation catalogue, see the [features reference](../features.md).

## Working with agents

- [Key Constraints & Sequencing](workflows.md) — what must happen before what
- [Behavioral Rules](behavioral-rules.md) — verification and destructive-operation safety
- [Anti-Patterns to Avoid](anti-patterns.md) — common mistakes and the correct approach
- [Gotchas & Known Limits](gotchas.md) — surprising Excel behaviour and workarounds
- [Agent Mode in Excel](agent-mode.md) — watching an agent drive the visible Excel window

## Workbooks, sheets and cells

- [Workbook Lifecycle](workbook.md)
- [Worksheet Operations](worksheet.md)
- [Ranges, Number Formats & Formatting](range.md)
- [Excel Tables](table.md)
- [Window Management](window.md)

## Data and the model

- [Power Query](powerquery.md)
- [M Code Syntax](m-code-syntax.md)
- [Data Model & DAX](datamodel.md)
- [DMV Query Reference](dmv-reference.md)
- [PivotTables](pivottable.md)
- [QueryTables](querytable.md)
- [What-If Analysis](analysis.md)
- [XML Maps](xmlmap.md)

## Visuals and output

- [Charts](chart.md)
- [Conditional Formatting](conditionalformat.md)
- [Slicers](slicer.md)
- [Drawing Objects](drawing.md)
- [Screenshots & Visual Verification](screenshot.md)
- [Dashboards & Reports](dashboard.md)
