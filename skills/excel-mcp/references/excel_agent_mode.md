# Agent Mode in Excel — Watch AI Work

Excel MCP's Agent Mode lets users watch AI operations happen in real-time. Instead of hidden automation, users see Excel respond to commands live — like Microsoft's Agent Mode in Excel, but with the full power of 225 COM-level operations.

> **MCP Server feature only.** Agent Mode uses conversational UI to ask about visibility preferences.

## When to Offer Agent Mode

**Always ask the user** at session start whether they want Excel visible or hidden. Present two clear choices using action cards:

> **Watch me work** — Show Excel side-by-side so you see every change live. Operations run slightly slower because Excel renders each update on screen.
>
> **Work in background** — Keep Excel hidden for maximum speed. You won't see changes until the task is done, but operations complete faster.

**Skip asking** only when the user has already stated a preference:
- User says "show me", "let me watch", "I want to see" → Show immediately
- User says "just do it", "work in background" → Keep hidden
- Simple one-shot operations (e.g., "what's in A1?") → Keep hidden, no need to ask

## Three Workflows

### 1. Agent Mode — Interactive Side-by-Side

User watches AI build a spreadsheet in real-time, side-by-side with the AI assistant.

```
1. file(open, path='report.xlsx')
2. window(show)                                    → Make visible
3. window(arrange, preset='right-half')            → Excel on right, AI on left
4. window(set-status-bar, text='Creating headers...') → Live feedback
5. range(set-values, ...)                          → User sees data appear
6. window(set-status-bar, text='Building PivotTable...')
7. pivottable(create, ...)                         → User sees PivotTable form
8. window(set-status-bar, text='Adding chart...')
9. chart(create-from-range, ...)                   → User sees chart render
10. window(clear-status-bar)                       → Clean up status bar
11. ASK: "I've finished the report. Would you like me to save and close, or keep it open?"
```

**Key behaviors:**
- Status bar shows what operation is in progress
- Ask before closing — user may want to inspect or make manual changes
- Narrate findings alongside visual changes

### 2. Presentation Mode — Guided Walkthrough

AI navigates through a completed workbook, explaining findings while the user watches.

```
1. file(open, path='analysis.xlsx')
2. window(show)
3. window(set-state, windowState='maximized')      → Full screen for best visibility
4. window(set-status-bar, text='Reviewing Sales sheet...')
5. "On the Sales sheet, you can see quarterly revenue trending up 15%..."
6. worksheet(activate, name='Analysis')            → Switch to next sheet
7. window(set-status-bar, text='Reviewing Analysis sheet...')
8. "The Analysis sheet shows the PivotTable breakdown by region..."
9. screenshot(capture-sheet)                       → Capture for AI context
10. window(clear-status-bar)
11. "Here's what I found: [summary with insights]"
```

**Key behaviors:**
- Maximized for full visibility
- Navigate sheets while narrating
- Screenshots at key points for AI context
- End with comprehensive summary

### 3. Debug Mode — Step-by-Step Inspection

AI performs operations one at a time, showing results between each step for troubleshooting.

```
1. file(open, path='broken.xlsx')
2. window(show)
3. window(arrange, preset='right-half')
4. window(set-status-bar, text='Inspecting Power Query...')
5. powerquery(list)
6. "Found 3 queries. Let me check each one..."
7. powerquery(view, queryName='Sales')
8. window(set-status-bar, text='Refreshing Sales query...')
9. powerquery(refresh, queryName='Sales')
10. screenshot(capture-sheet)                      → Show result
11. "Sales query refreshed successfully. 150 rows loaded. Moving to next..."
12. window(set-status-bar, text='Refreshing Products query...')
13. powerquery(refresh, queryName='Products')
14. "Products query failed: [error]. Let me fix the M code..."
15. window(clear-status-bar)
```

**Key behaviors:**
- Pause between operations to show intermediate state
- Screenshot after each significant step
- Narrate what's happening and what was found
- Ideal for diagnosing query failures, formula errors, data issues

## Status Bar Best Practices

Use `window(set-status-bar)` to show operation progress in Excel's status bar:

| Operation | Status Bar Text |
|-----------|----------------|
| Writing data | `"ExcelMcp: Writing 500 rows to Sales sheet..."` |
| Building PivotTable | `"ExcelMcp: Building PivotTable from Sales data..."` |
| Refreshing query | `"ExcelMcp: Refreshing Power Query 'Revenue'..."` |
| Creating chart | `"ExcelMcp: Creating bar chart from sales data..."` |
| Formatting | `"ExcelMcp: Applying currency formatting..."` |

**Always clear** with `window(clear-status-bar)` when the workflow completes.

**Only set when visible**: Status bar text is only useful when Excel is visible. Skip status bar calls when Excel is hidden.

## Asking About Visibility

When starting a session, present the visibility choice as action cards so the user can pick with one click:

> **Watch me work** — Show Excel side-by-side so you see every change live. Operations run slightly slower because Excel renders each update on screen.
>
> **Work in background** — Keep Excel hidden for maximum speed. You won't see changes until the task is done, but operations complete faster.

If the user picks "Watch me work":
1. `window(show)` → Make Excel visible
2. `window(arrange, preset='right-half')` → Position for side-by-side
3. Use `window(set-status-bar)` throughout the workflow for live progress

If the user picks "Work in background" or doesn't respond, keep Excel hidden and skip all status bar calls.
