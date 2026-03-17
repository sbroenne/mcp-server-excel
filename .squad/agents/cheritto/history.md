# Cheritto — History

## Core Context

- **Project:** A Windows COM interop MCP server and CLI for programmatic Excel automation with 25 tools and 225 operations.
- **Role:** Platform Dev
- **Joined:** 2026-03-15T10:42:22.620Z

## Learnings

<!-- Append learnings below -->
- 2026-03-16: Bug 3 is already supported under `range.set-number-format`; the actual gap is that the formatting surface is split between `range` (display formats) and `range_format` (visual styling), and top-level feature tables currently hide that split.
- 2026-03-16: Bug 5 is already supported under `range_format.auto-fit-columns` / `auto-fit-rows`; the main weakness is discoverability, not backend capability.
- 2026-03-16: Bug 4 is a real surface gap. A `List<...>` batch request is consistent with existing patterns such as `IRangeEditCommands.Sort(... List<SortColumn> ...)`, so the likely risk is in API review and docs parity, not generator feasibility.
