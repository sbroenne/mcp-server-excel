# Project Context

- **Owner:** sbroenne
- **Project:** ExcelMcp (sbroenne/mcp-server-excel) — Windows-only automation of installed desktop Excel through COM; the MCP Server and `excelcli` are equal entry points over shared Core commands
- **Stack:** C#/.NET (SDK from `global.json`), Excel COM, MCP SDK, generated Service/CLI/MCP surfaces, PowerShell scripts, MkDocs website, VS Code extension (TypeScript)
- **Created:** 2026-10-06T17:10:20.668+02:00

## Core Context

Quality Engineer owns `tests/` (see `tests/AGENTS.md`), regression evidence, and local `scripts\Test-E2E.ps1`. Desktop Excel is required for COM tests and every Excel test run is sequential. `llm-tests/` runs only when explicitly requested.

## Recent Updates

📌 Team initialized on 2026-10-06T17:10:20.668+02:00

## Learnings

<!-- Append new learnings below. Each entry is something lasting about the project. -->
