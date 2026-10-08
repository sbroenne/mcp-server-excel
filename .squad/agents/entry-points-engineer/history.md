# Project Context

- **Owner:** sbroenne
- **Project:** ExcelMcp (sbroenne/mcp-server-excel) — Windows-only automation of installed desktop Excel through COM; the MCP Server and `excelcli` are equal entry points over shared Core commands
- **Stack:** C#/.NET (SDK from `global.json`), Excel COM, MCP SDK, generated Service/CLI/MCP surfaces, PowerShell scripts, MkDocs website, VS Code extension (TypeScript)
- **Created:** 2026-10-06T17:10:20.668+02:00

## Core Context

Entry Points Engineer owns Service, CLI, MCP Server, and the generators. Core `[ServiceCategory]` interfaces drive generated routing, so change contracts or generators, never emitted code. Follow `docs/agents/rules/coverage-prevention-strategy.md` and `docs/agents/rules/mcp-server-guide.md`.

## Recent Updates

📌 Team initialized on 2026-10-06T17:10:20.668+02:00

## Learnings

<!-- Append new learnings below. Each entry is something lasting about the project. -->
