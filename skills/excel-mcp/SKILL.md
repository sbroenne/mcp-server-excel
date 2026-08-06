---
name: excel-mcp
description: >
  Excel MCP Server skill for Windows workbook automation. Use when an assistant
  needs rich MCP tools to create, inspect, modify, format, or analyze Excel files.
  Supports Power Query (M), Data Model/DAX, PivotTables, Tables, Ranges, Charts,
  Slicers, formatting, screenshots, VBA macros, connections, and calculation mode.
  Triggers: Excel, spreadsheet, workbook, xlsx, xlsm, Power Query, DAX, PivotTable,
  chart, dashboard, VBA, MCP.
---

# Excel MCP Server Skill

Provides 246 Excel operations via Model Context Protocol. The MCP Server hosts the ExcelMCP Service in-process and calls it directly for low-latency Excel automation. Tools are auto-discovered - this documents quirks, workflows, and gotchas.

## Copilot compact profile

The Copilot plugin starts the `copilot-compact` profile: 9 tools (`file`, `workflow`, `worksheet`, `range`, `range_edit`, `range_format`, `worksheet_style`, `layout`, and `calculation_mode`). The full profile is 29 tools/246 operations; compact intentionally omits those domain tools and adds the `layout` facade for deterministic report formatting and outlines. Treat live `tools/list` and `workflow(capabilities)` as authoritative.

In compact mode, use `workflow(open-and-describe)` to open and inspect, then `workflow(execute-plan)` for two or more compatible edits. A plan may request one checkpoint (`checkpoint_mode: once`); it is ordered, non-atomic, and stops at the first error. Prefer the bounded verification receipt returned by the workflow or mutation, scoped to the reported ranges, instead of broad readbacks. If the receipt is partial or layout is visual, use the targeted range or screenshot tool that is available.

Reuse an idempotency key only when the prior result is known and the command is unchanged. A timeout, cancellation, process death, or connection loss after dispatch has an unknown outcome—inspect the workbook or recovery evidence before retrying. Keep MCP session IDs separate from CLI session IDs. On failure, inspect the failed plan index and receipt, recover from the checkpoint when configured, and close with an explicit save choice.

## Workflow Checklist

| Step | Tool | Action | When |
|------|------|--------|------|
| 1. Confirm runtime | `workflow` | `capabilities` | Once, only when `workflow` is in `tools/list` |
| 2. Open and inspect | `workflow` or `file` | `open-and-describe`, else `open` | Start of workbook work |
| 3. Create sheets | `worksheet` | `create`, `rename` | If needed |
| 4. Write and format | `workflow` or domain tools | `execute-plan`, or direct actions | Batch 2+ compatible edits; use direct tools for one-offs |
| 5. Structure | `table` | `create` | Convert tabular data to tables |
| 6. Verify | relevant read tool or `screenshot` | targeted readback | Before claiming completion |
| 7. Save & close | `file` | `close` with explicit `save` | Always last |

## Copilot MCP Runtime Truth

Treat the current MCP `tools/list` response as authoritative. Repository source, this skill's version, and an installed plugin do not prove which binary the client loaded. Use `workflow` only when it appears in `tools/list`; then call `workflow(action: 'capabilities')` once and record `workflowInterfaceVersion`, `serverVersion`, `buildFingerprint`, and `toolProfile`. If it is absent, use `file` plus the domain tools.

When available, `workflow(action: 'open-and-describe')` opens a workbook and returns a fresh, bounded workbook manifest plus a session ID in one call. It replaces the common open/list/used-range discovery sequence. The returned session is still open and must later be closed explicitly.

Use `workflow(action: 'execute-plan')` for two or more compatible ordered commands. Each item contains a service `command` such as `range.set-values` and an `args` object keyed by that command's service parameter names. Use direct domain tools when the command arguments are uncertain or the task is a single operation.

`execute-plan` is ordered and stops on the first error by default. It is not atomic, does not roll back earlier successful steps, and does not include save/close. Inspect its results and failed index, verify the workbook, then call `file(action: 'close')` with an explicit save choice.

## Safety, Retry, and Session Discipline

- Use review, checkpoint, journal, verification, and idempotency features only when the loaded schema exposes them and the session is configured for them.
- `review_only` is a machine-readable review step. A review ID does not prove a human approved the action; use a genuine host/user interaction when human consent is required.
- A timeout, cancellation, Excel process death, or connection loss after dispatch has an unknown outcome. Never blindly replay it, even with an idempotency key; inspect the workbook or recovery evidence first.
- Trust a verification result only for its exact reported scope and status. Read back cells/formulas, or use a screenshot when visual layout matters, for anything not fully covered.
- MCP and CLI sessions are separate. Never pass a session ID from one interface to the other.

## Preconditions

- Windows host with Microsoft Excel installed (2016+)
- Use full Windows paths: `C:\Users\Name\Documents\Report.xlsx`
- Excel files must not be open in another Excel instance

## Calculation Mode Workflow (Batch Performance)

Use `calculation_mode` for **bulk write performance optimization**. When writing many values or formulas, disable auto-recalc to avoid recalculating after every cell:

```
1. calculation_mode(action: 'set-mode', mode: 'manual')  → Disable auto-recalc
2. Perform all writes (range set-values, set-formulas)
3. calculation_mode(action: 'calculate', scope: 'workbook')  → Recalculate once
4. calculation_mode(action: 'set-mode', mode: 'automatic')  → Restore default
```

**Note:** You do NOT need manual mode to read formulas - `range get-formulas` returns formula text regardless of calculation mode.

## CRITICAL: Execution Rules (MUST FOLLOW)

### Rule 1: Discover Before Asking; Ask When Required

Discover active sessions and workbook structure before asking questions the tools can answer. Ask when a required workbook path, irreversible choice, credential step, or genuine human approval is missing; never guess those inputs.

| Bad (Asking) | Good (Discovering) |
|--------------|-------------------|
| "Which Excel file should I use?" | `file(list)` → use the unambiguous open session; otherwise ask for a path |
| "What's the table name?" | `table(list)` → discover tables |
| "Which sheet has the data?" | `worksheet(list)` → check all sheets |
| "Should I create a PivotTable?" | YES - create it on a new sheet |

Use discovery tools for workbook facts; reserve questions for information or authority the tools cannot supply.

### Rule 2: Always End With a Text Summary

**NEVER end your turn with only a tool call.** After completing all operations, always provide a brief text message confirming what was done. Silent tool-call-only responses are incomplete.

### Rule 3: Format Data Professionally

Always apply number formats after setting values:

| Data Type | Format Code | Result |
|-----------|-------------|--------|
| USD | `$#,##0.00` | $1,234.56 |
| EUR | `€#,##0.00` | €1,234.56 |
| Percent | `0.00%` | 15.00% |
| Date (ISO) | `yyyy-mm-dd` | 2025-01-22 |

**Workflow:**
```
1. range set-values (data is now in cells)
2. range set-number-format (apply format)
```

### Rule 4: Use Excel Tables (Not Plain Ranges)

Always convert tabular data to Excel Tables:

```
1. range set-values (write data including headers)
2. table create tableName="SalesData" rangeAddress="A1:D100"
```

**Why:** Structured references, auto-expand, required for Data Model/DAX.

### Rule 5: Session Lifecycle

```
1. file(action: 'open', path: '...')  → sessionId
2. All operations use sessionId
3. file(action: 'close', save: true)  → saves and closes
```

**Unclosed sessions leave Excel processes running, locking files.**

### Rule 6: Data Model Prerequisites

DAX operations require tables in the Data Model:

```
Step 1: Create table → Table exists
Step 2: table(action: 'add-to-datamodel') → Table in Data Model
Step 3: datamodel(action: 'create-measure') → NOW this works
```

### Rule 7: Power Query Development Lifecycle

**BEST PRACTICE: Test-First Workflow**

```
1. powerquery(action: 'evaluate', mCode: '...') → Test WITHOUT persisting
2. powerquery(action: 'create', ...) → Store validated query
3. powerquery(action: 'refresh', ...) → Load data
```

**Why evaluate first:**
- Catches syntax errors and missing sources BEFORE creating permanent queries
- Better error messages than COM exceptions from create/update
- See actual data preview (columns + sample rows)
- No cleanup needed - like a REPL for M code
- Skip only for trivial literal tables

**Common mistake:** Creating/updating without evaluate → pollutes workbook with broken queries

### Rule 8: Targeted Updates Over Delete-Rebuild

- **Prefer**: `set-values` on specific range (e.g., `A5:C5` for row 5)
- **Avoid**: Deleting and recreating entire structures

**Why:** Preserves formatting, formulas, and references.

### Rule 9: Follow suggestedNextActions

Error responses include actionable hints:
```json
{
  "success": false,
  "errorMessage": "Table 'Sales' not found in Data Model",
  "suggestedNextActions": ["table(action: 'add-to-data-model', tableName: 'Sales')"]
}
```

## Tool Selection Quick Reference

| Task | Tool | Key Action |
|------|------|------------|
| Create/open/save workbooks | `file` | open, create, close |
| Confirm loaded optimized surface | `workflow` | capabilities |
| Open and inspect in one call | `workflow` | open-and-describe |
| Run compatible ordered edits | `workflow` | execute-plan |
| Write/read cell data | `range` | set-values, get-values |
| Format cells | `range` | set-number-format |
| Create tables from data | `table` | create |
| Add table to Power Pivot | `table` | add-to-data-model |
| Create DAX formulas | `datamodel` | create-measure |
| Create PivotTables | `pivottable` | create, create-from-datamodel |
| Filter with slicers | `slicer` | set-slicer-selection |
| Create charts | `chart` | create-from-range |
| Control calculation mode | `calculation_mode` | get-mode, set-mode, calculate |
| Visual verification | `screenshot` | capture, capture-sheet |

## Reference Documentation

See `references/` for detailed guidance:

- [Core execution rules and LLM guidelines](./references/behavioral-rules.md)
- [Common mistakes to avoid](./references/anti-patterns.md)
- [Bulk write performance optimization](./references/calculation.md)
- [Data Model constraints and patterns](./references/workflows.md)
- [Charts and formatting](./references/chart.md)
- [Conditional formatting operations](./references/conditionalformat.md)
- [Dashboard and report best practices](./references/dashboard.md)
- [Data Model/DAX specifics](./references/datamodel.md)
- [DMV query reference for Data Model analysis](./references/dmv-reference.md)
- [Excel agent mode and advanced automation](./references/excel_agent_mode.md)
- [Gotchas and known limits](./references/gotchas.md)
- [Power Query M code syntax reference](./references/m-code-syntax.md)
- [PivotTable operations](./references/pivottable.md)
- [Power Query specifics](./references/powerquery.md)
- [Range operations and number formats](./references/range.md)
- [Screenshot and visual verification](./references/screenshot.md)
- [Slicer operations](./references/slicer.md)
- [Table operations](./references/table.md)
- [Window and visibility operations](./references/window.md)
- [Worksheet operations](./references/worksheet.md)
