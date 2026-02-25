# excel-mcp-skill

An [Agent Skill](https://agentskills.io) for automating Microsoft Excel via the [Excel MCP Server](https://excelmcpserver.dev).

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent how to automate Excel through 225 MCP operations:

- **Workbook management** — open, create, save, close
- **Range operations** — read/write values, formatting, formulas
- **Tables & PivotTables** — create, modify, refresh
- **Charts** — create and configure chart types
- **Power Query (M code)** — create and edit queries
- **Data Model (DAX)** — add measures and calculated columns
- **Conditional formatting, slicers, VBA macros**, and more

## Requirements

- Windows with Microsoft Excel 2016+ installed
- [Excel MCP Server](https://github.com/sbroenne/mcp-server-excel) running

## Install

```bash
npx skillpm install excel-mcp-skill
```

Or with npm directly:

```bash
npm install excel-mcp-skill
```

## License

MIT
