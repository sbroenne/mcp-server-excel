# McCauley — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with 25 tools and 225 operations.
- **Role:** Lead
- **Joined:** 2026-03-15T10:42:22.614Z

## Learnings

<!-- Append learnings below -->
- 2026-03-16: Treat reported `set-values` width failures above 13 columns as MCP/service/client-shape suspects first, not automatic Core defects; Core integration coverage already exercises `SetValues` on `A1:P1` successfully.
- 2026-03-16: Split "missing formatting capability" reports into real product gaps versus discoverability gaps. `set-number-format` already exists on `range`, and `auto-fit-columns` already exists on `range_format`.
- 2026-03-16: Do not accept "activate the new sheet" as the default fix for post-create write issues without proof; `SetValues` resolves by explicit sheet name, so a non-`A1` write failure after sheet creation needs a targeted reproduction test before changing lifecycle behavior.
- 2026-03-17: Final gate for Bug 1 stays green when the fix validates rectangular row widths at the command boundary, mirrors the same guard in `SetFormulas`, and keeps wider named-range or multi-area behavior explicitly out of scope until a separate repro exists.
