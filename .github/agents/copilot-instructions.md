# mcp-server-excel Development Guidelines

Auto-generated from all feature plans. Last updated: 2025-12-19

## Active Technologies
- C# 14 / .NET 10.0 + ModelContextProtocol, Microsoft.Extensions.*, Application Insights (006-dotnet10-upgrade)
- N/A (Excel files managed via COM) (006-dotnet10-upgrade)
- Markdown/YAML (documentation only, no code) + None (static files following agentskills.io spec) (007-agent-skills)
- N/A (file-based skill package) (007-agent-skills)
- Markdown/YAML (documentation-only feature, no .NET code changes) + agentskills.io specification (YAML frontmatter + Markdown body) (007-agent-skills)
- File-based (`skills/excel-mcp/` directory at repo root) (007-agent-skills)

- C# / .NET 10 + Excel COM automation via `dynamic` + `ExcelMcp.ComInterop`, MCP SDK (`ModelContextProtocol`), `System.Text.Json`, CLI via `Spectre.Console.Cli` (001-rename-queries-tables)

## Project Structure

```text
src/
tests/
```

## Commands

# Add commands for C# / .NET 10

## Code Style

C# / .NET 10: Follow standard conventions

## Recent Changes
- 007-agent-skills: Added Markdown/YAML (documentation-only feature, no .NET code changes) + agentskills.io specification (YAML frontmatter + Markdown body)
- 007-agent-skills: Added Markdown/YAML (documentation-only feature, no .NET code changes) + agentskills.io specification (YAML frontmatter + Markdown body)
- 007-agent-skills: Added Markdown/YAML (documentation only, no code) + None (static files following agentskills.io spec)


<!-- MANUAL ADDITIONS START -->
<!-- MANUAL ADDITIONS END -->
