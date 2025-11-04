# GitHub Copilot Instructions - ExcelMcp

> **🎯 Optimized for AI Coding Agents** - Modular, path-specific instructions following GitHub Copilot best practices

## 📋 Quick Navigation

**Start here** → Read [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) first (14 mandatory rules)

**Path-Specific Instructions** (auto-applied based on file context):
- 🧪 [Testing Strategy](instructions/testing-strategy.instructions.md) - Test templates, essential patterns
- 📊 [Excel COM Interop](instructions/excel-com-interop.instructions.md) - COM patterns, cleanup
- 🔌 [Excel Connection Types](instructions/excel-connection-types-guide.instructions.md) - Connection types, COM API
- 🏗️ [Architecture Patterns](instructions/architecture-patterns.instructions.md) - Command pattern, batch management
- 🧠 [MCP Server Guide](instructions/mcp-server-guide.instructions.md) - MCP tools, protocol
- 🔄 [Development Workflow](instructions/development-workflow.instructions.md) - PR process, CI/CD
- 🐛 [Bug Fixing Checklist](instructions/bug-fixing-checklist.instructions.md) - 6-step bug fix process
- 📚 [README Management](instructions/readme-management.instructions.md) - Documentation quick reference
- 📁 [Documentation Structure](instructions/documentation-structure.instructions.md) - Where to put docs, avoid temporary files

---

## What is ExcelMcp?

**ExcelMcp** is a Windows-only toolset for programmatic Excel automation via COM interop, designed for coding agents and automation scripts.

**Four Layers:**
1. **ComInterop** (`src/ExcelMcp.ComInterop`) - Reusable COM automation patterns (STA threading, session management, batch operations, OLE message filter)
2. **Core** (`src/ExcelMcp.Core`) - Excel-specific business logic (Power Query, VBA, worksheets, parameters)
3. **CLI** (`src/ExcelMcp.CLI`) - Command-line interface for scripting
4. **MCP Server** (`src/ExcelMcp.McpServer`) - Model Context Protocol for AI assistants

**Key Capabilities:**
- **Range Operations** (Phase 1 implementation in progress) - Unified API for all range data operations (get/set values/formulas, clear variants, find/replace, sort, insert/delete, copy/paste, UsedRange, CurrentRegion, hyperlinks)
- Power Query M code management (import, export, update, refresh)
- VBA macro management (list, import, export, run)
- Worksheet lifecycle management (list, create, rename, copy, delete)
- Named range parameters (create, delete, update, list, get/set single values)
- Data Model operations (list tables/measures/relationships, export measures, refresh, delete)
- Connection management (list, view, import/export, update, refresh, test, properties)

---

## 🎯 Development Quick Start

### Common Tasks
- **Add new command** → Follow patterns in [Architecture Patterns](instructions/architecture-patterns.instructions.md)
- **Excel COM work** → Reference [Excel COM Interop](instructions/excel-com-interop.instructions.md)
- **Modify session/batch code** → MUST run OnDemand tests (see [CRITICAL-RULES.md](instructions/critical-rules.instructions.md))
- **Add MCP tool** → Follow [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- **Create PR** → Follow [Development Workflow](instructions/development-workflow.instructions.md)
- **Fix bug** → Use [Bug Fixing Checklist](instructions/bug-fixing-checklist.instructions.md) (6-step process)
- **Add documentation** → Use [Documentation Structure](instructions/documentation-structure.instructions.md) (avoid temporary files)
- **Migrate tests to batch API** → See BATCH-API-MIGRATION-PLAN.md for comprehensive guide
- **Create simple tests** → Use ConnectionCommandsSimpleTests.cs or SetupCommandsSimpleTests.cs as template
- **Range API implementation** → See [Range API Specification](../specs/RANGE-API-SPECIFICATION.md) for complete design (38 methods, MCP-first, breaking changes acceptable)

### Test Execution
```bash
# Development (fast feedback - excludes VBA)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Pre-commit (requires Excel - excludes VBA)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Session/batch cleanup (MANDATORY when modifying session/batch code)
dotnet test --filter "RunType=OnDemand"

# VBA tests (manual only - requires VBA trust enabled)
dotnet test --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"
```

### Batch API Pattern (Current Standard)
```csharp
// Core Commands - Always use batch parameter
public async Task<OperationResult> MethodAsync(ExcelBatch batch, string arg1)
{
    // batch.Book gives access to workbook
    // batch.FilePath has the file path
    return new OperationResult { Success = true };
}

// CLI Commands - Wrap in try-catch
public int Method(string[] args)
{
    ResultType result;
    try
    {
        var task = Task.Run(async () =>
        {
            await using var batch = await ExcelSession.BeginBatchAsync(filePath);
            var opResult = await _coreCommands.MethodAsync(batch, arg1);
            await batch.SaveAsync(); // if changes made
            return opResult;
        });
        result = task.GetAwaiter().GetResult();
    }
    catch (Exception ex)
    {
        AnsiConsole.MarkupLine($"[red]Error:[/] {ex.Message.EscapeMarkup()}");
        return 1;
    }
    
    if (result.Success) { /* format output */ return 0; }
    else { /* show error */ return 1; }
}

// Tests - Use batch API
[Fact]
public async Task TestMethod()
{
    await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
    var result = await _commands.MethodAsync(batch, args);
    Assert.True(result.Success);
}
```

---

## 📎 Related Resources

**For Excel automation in other projects:**
- Copy `docs/excel-powerquery-vba-copilot-instructions.md` to your project's `.github/copilot-instructions.md`

**Project Documentation:**
- [Commands Reference](../docs/COMMANDS.md)
- [Installation Guide](../docs/INSTALLATION.md)
- [Range API Specification](../specs/RANGE-API-SPECIFICATION.md) - Comprehensive design for unified range operations (Phase 1 implementation)
- [Range Refactoring Analysis](../specs/RANGE-REFACTORING-ANALYSIS.md) - LLM perspective on consolidating fragmented commands

---

## 🔄 Continuous Learning

After completing significant tasks, update these instructions with lessons learned. See [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) Rule 4.

### Key Lessons Learned

**COM Interop Extraction:** ExcelMcp.ComInterop is separate reusable library. Tests with Excel process side effects use `[Trait("RunType", "OnDemand")]`.

**Batch API Migration:** Create NEW simple tests instead of converting complex old tests. CLI commands using `BeginBatchAsync` need try-catch wrapping.

**Excel Type 3/4 Handling:** Excel returns type 4 (WEB) for TEXT connections. Handle BOTH types 3 and 4 in all connection property methods.

**MCP Prompt Design:** Prompts should be SHORT user shortcuts, not tutorials. Domain knowledge only - LLMs already know programming languages.

**Range API Design:** Single cell = 1x1 range. COM-backed only. MCP-first implementation. Breaking changes acceptable during active development.

**Refactoring Strategy:** File recreation > incremental edits when removing 50%+ content. Use partial classes for 500+ line Core commands, single file OK for MCP tools up to 1400 lines.

**CLI Testing:** Don't duplicate integration tests - CLI only tests argument parsing, exit codes, CSV conversion.

**Spec Validation:** Always search Microsoft official docs FIRST using mcp_microsoft_doc tools. Never trust secondary sources or assumptions.

**QueryTable Persistence:** `RefreshAll()` is async - doesn't persist QueryTables. Use individual `queryTable.Refresh(false)` synchronously.

**VS Code Extensions:** Use kebab-case IDs, validate security compliance, test installation readiness, maintain consistent IDs across package.json and TypeScript.

**Bulk Refactoring:** Use `replace_string_in_file` with 3-5 lines context. Use VS Code built-in tools: grep_search, semantic_search, file_search, get_errors. Avoid PowerShell for code changes.

**Tool Selection Priority:** Code changes → `replace_string_in_file` | File creation → `create_file` | Find code → `grep_search`/`semantic_search` | Check errors → `get_errors` | Build/test/git → `run_in_terminal`

**Pre-Commit:** Search for TODO/FIXME/HACK markers, resolve all, delete commented-out code, verify tests pass, update docs if behavior changed.

---

## 📚 How Path-Specific Instructions Work

GitHub Copilot automatically loads instructions based on the files you're working with:

- Working in `tests/**/*.cs`? → [Testing Strategy](instructions/testing-strategy.instructions.md) auto-applies
- Working in `src/ExcelMcp.Core/**/*.cs`? → [Excel COM Interop](instructions/excel-com-interop.instructions.md) auto-applies
- Working in `src/ExcelMcp.ComInterop/**/*.cs`? → Low-level COM utilities (minimal dependencies)
- Working in `src/ExcelMcp.McpServer/**/*.cs`? → [MCP Server Guide](instructions/mcp-server-guide.instructions.md) auto-applies
- Working in `.github/workflows/**/*.yml`? → [Development Workflow](instructions/development-workflow.instructions.md) auto-applies
- **All files** → [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) always applies

This modular approach ensures you get relevant context without overwhelming the AI with unnecessary information.

