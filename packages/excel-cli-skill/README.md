# excel-cli-skill

An [Agent Skill](https://agentskills.io) for automating Microsoft Excel via the [excelcli](https://excelmcpserver.dev) command-line tool.

## What this skill does

When loaded by an AI agent (Claude, Codex, Cursor, Gemini CLI, etc.), this skill teaches the agent how to automate Excel from scripts and CI/CD pipelines:

- **Workbook management** — open, create, save, close
- **Range operations** — read/write values, formatting, formulas
- **Tables & PivotTables** — create, modify, refresh
- **Charts** — create and configure chart types
- **Power Query (M code)** — create and edit queries
- **Data Model (DAX)** — add measures and calculated columns
- **VBA macros, conditional formatting**, and more

## Requirements

- Windows with Microsoft Excel 2016+ installed
- Install the CLI: `dotnet tool install --global Sbroenne.ExcelMcp.CLI`

## Install

```bash
npx skillpm install excel-cli-skill
```

Or with npm directly:

```bash
npm install excel-cli-skill
```

## License

MIT
