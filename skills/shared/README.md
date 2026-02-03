# Skills Shared References

This directory contains behavioral guidance and tool-specific quirks documentation that is shared between CLI and MCP skill packages.

## Source of Truth

**IMPORTANT:** The files in this directory are synchronized from the MCP Server prompts:

```
Source: src/ExcelMcp.McpServer/Prompts/Content/
Target: skills/shared/
```

## File Categories

### 1. Excel Tool-Specific Guidance (Synced from MCP Server)

These files are embedded in the MCP Server and served via the MCP protocol. They are copied here for agent skills distribution:

- `excel_chart.md` - Chart operations and PivotCharts
- `excel_conditionalformat.md` - Conditional formatting rules
- `excel_datamodel.md` - Data Model (Power Pivot) and DAX
- `excel_powerquery.md` - Power Query M code workflows
- `excel_range.md` - Range operations and number formatting
- `excel_slicer.md` - Slicer operations for filtering
- `excel_table.md` - Excel Table operations
- `excel_worksheet.md` - Worksheet lifecycle

### 2. Agent Skill-Specific Guidance (Skills Only)

These files are unique to agent skills and not part of the MCP Server prompts:

- `anti-patterns.md` - Common mistakes to avoid
- `behavioral-rules.md` - Core execution rules for LLMs
- `workflows.md` - Data Model constraints and patterns

## Keeping Files in Sync

### For Maintainers

When updating tool guidance:

1. **Edit the MCP Server prompts first**: `src/ExcelMcp.McpServer/Prompts/Content/*.md`
2. **Copy to skills/shared**: Run the sync script (see below)
3. **Test both locations**: Verify MCP prompts and skill packages

### Sync Script

```powershell
# Copy MCP prompts to skills/shared (overwrite)
$files = @(
    "excel_chart.md",
    "excel_conditionalformat.md",
    "excel_datamodel.md",
    "excel_powerquery.md",
    "excel_range.md",
    "excel_slicer.md",
    "excel_table.md",
    "excel_worksheet.md"
)

foreach ($file in $files) {
    Copy-Item -Path "src/ExcelMcp.McpServer/Prompts/Content/$file" `
              -Destination "skills/shared/$file" `
              -Force
}
```

### Verification

```powershell
# Verify all files are synchronized
$files = @("excel_chart.md", "excel_conditionalformat.md", "excel_datamodel.md", 
           "excel_powerquery.md", "excel_range.md", "excel_slicer.md", 
           "excel_table.md", "excel_worksheet.md")

foreach ($file in $files) {
    $mcp = "src/ExcelMcp.McpServer/Prompts/Content/$file"
    $shared = "skills/shared/$file"
    if ((Get-FileHash $mcp).Hash -eq (Get-FileHash $shared).Hash) {
        Write-Host "✓ $file - SYNCHRONIZED" -ForegroundColor Green
    } else {
        Write-Host "✗ $file - OUT OF SYNC" -ForegroundColor Red
    }
}
```

## Why This Architecture?

### MCP Server Prompts are Source of Truth

- **Embedded in the server**: Prompts are compiled as embedded resources
- **Served via MCP protocol**: LLMs can request prompts via `prompts/list` and `prompts/get`
- **Actively maintained**: Updated alongside tool implementations
- **Server-specific**: Documents actual MCP Server behavior

### Agent Skills Need Same Content

- **CLI and MCP share concepts**: Power Query, Data Model, Tables work identically
- **Different syntax, same guidance**: `excelcli -q powerquery list` vs `excel_powerquery(action: 'list')`
- **Distributed in ZIP packages**: Agent skills are packaged for Copilot/Claude/Cursor
- **Offline availability**: Agents don't have access to MCP prompts during development

## Common Questions

**Q: Why duplicate the files?**
A: MCP Server prompts are embedded resources (compiled). Agent skills need standalone markdown files for distribution.

**Q: Should CLI and MCP have different guidance?**
A: No! The Excel concepts are identical. Only the interface differs (CLI commands vs MCP tools).

**Q: What if I forget to sync?**
A: Files will diverge. Always edit MCP prompts first, then copy to skills/shared.

**Q: Can I edit skills/shared directly?**
A: No! Changes will be overwritten on next sync. Edit MCP prompts first.
