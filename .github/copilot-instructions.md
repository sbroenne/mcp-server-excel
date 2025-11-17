# GitHub Copilot Instructions - ExcelMcp

> **🎯 Optimized for AI Coding Agents** - Modular, path-specific instructions

## 📋 Critical Files (Read These First)

**ALWAYS read when working on code:**
- [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) - 19 mandatory rules (Success flag, COM cleanup, tests, etc.)
- [Architecture Patterns](instructions/architecture-patterns.instructions.md) - Batch API, command pattern, resource management

**Read based on task type:**
- Adding/fixing commands → [Excel COM Interop](instructions/excel-com-interop.instructions.md)
- Writing tests → [Testing Strategy](instructions/testing-strategy.instructions.md)
- MCP Server work → [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- Creating PR → [Development Workflow](instructions/development-workflow.instructions.md)
- Fixing bugs → [Bug Fixing Checklist](instructions/bug-fixing-checklist.instructions.md)

**Less frequently needed:**
- [Excel Connection Types](instructions/excel-connection-types-guide.instructions.md) - Only for connection-specific work
- [README Management](instructions/readme-management.instructions.md) - Only when updating READMEs
- [Documentation Structure](instructions/documentation-structure.instructions.md) - Only when creating docs

---

## What is ExcelMcp?

**ExcelMcp** is a Windows-only toolset for programmatic Excel automation via COM interop, designed for coding agents and automation scripts.

**Four Layers:**
1. **ComInterop** (`src/ExcelMcp.ComInterop`) - Reusable COM automation patterns (STA threading, session management, batch operations, OLE message filter)
2. **Core** (`src/ExcelMcp.Core`) - Excel-specific business logic (Power Query, VBA, worksheets, parameters)
3. **CLI** (`src/ExcelMcp.CLI`) - Command-line interface for scripting
4. **MCP Server** (`src/ExcelMcp.McpServer`) - Model Context Protocol for AI assistants

---

## 🎯 Quick Reference

### Test Commands
```bash
# Fast feedback (excludes VBA)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Session/batch changes (MANDATORY)
dotnet test --filter "RunType=OnDemand"
```

### Code Patterns
```csharp
// Core: Always use batch parameter
public async Task<OperationResult> MethodAsync(IExcelBatch batch, string arg1)
{
    return await batch.Execute((ctx, ct) => {
        // Use ctx.Book for workbook access
        return ValueTask.FromResult(new OperationResult { Success = true });
    });
}

// CLI: Wrap Core calls
public int Method(string[] args)
{
    try {
        var task = Task.Run(async () => {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            return await _coreCommands.MethodAsync(batch, arg1);
        });
        var result = task.GetAwaiter().GetResult();
        return result.Success ? 0 : 1;
    } catch (Exception ex) {
        AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
        return 1;
    }
}

// Tests: Use batch API
[Fact]
public async Task TestMethod()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    var result = await _commands.MethodAsync(batch, args);
    Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
}
```

### Tool Selection
- Code changes → `replace_string_in_file` (3-5 lines context)
- Find code → `grep_search` or `semantic_search`
- Check errors → `get_errors`
- Build/test/git → `run_in_terminal`

---

## 🔄 Key Lessons (Update After Major Work)

**Success Flag:** NEVER `Success = true` with `ErrorMessage`. Set Success in try block, always false in catch.

**Batch API:** Create NEW simple tests. CLI needs try-catch wrapping.

**Excel Quirks:** Type 3/4 both handle TEXT. `RefreshAll()` unreliable. Use `queryTable.Refresh(false)`.

**MCP Design:** Prompts are shortcuts, not tutorials. LLMs know Excel/programming.

**Tool Priority:** `replace_string_in_file` > `grep_search` > `run_in_terminal`. Avoid PowerShell for code.

**Pre-Commit:** Search TODO/FIXME/HACK, delete commented code, verify tests, check docs.

**PR Review:** Check automated comments immediately (Copilot, GitHub Security). Fix before human review.

---

## 📚 How Path-Specific Instructions Work

GitHub Copilot auto-loads instructions based on files you're editing:

- `tests/**/*.cs` → [Testing Strategy](instructions/testing-strategy.instructions.md)
- `src/ExcelMcp.Core/**/*.cs` → [Excel COM Interop](instructions/excel-com-interop.instructions.md)
- `src/ExcelMcp.McpServer/**/*.cs` → [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- `.github/workflows/**/*.yml` → [Development Workflow](instructions/development-workflow.instructions.md)
- `**` (all files) → [CRITICAL-RULES.md](instructions/critical-rules.instructions.md)

Modular approach = relevant context without overload.

