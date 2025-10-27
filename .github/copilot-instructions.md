# GitHub Copilot Instructions - ExcelMcp

> **🎯 Optimized for AI Coding Agents** - Modular, path-specific instructions following GitHub Copilot best practices

## 📋 Quick Navigation

**Start here** → Read [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) first (5 mandatory rules)

**Path-Specific Instructions** (auto-applied based on file context):
- 🧪 [Testing Strategy](instructions/testing-strategy.instructions.md) - Test architecture, OnDemand pattern, filtering
- 📊 [Excel COM Interop](instructions/excel-com-interop.instructions.md) - COM patterns, cleanup, best practices
- 🔌 [Excel Connection Types](instructions/excel-connection-types-guide.instructions.md) - Connection types, COM API limitations, testing strategies
- 🏗️ [Architecture Patterns](instructions/architecture-patterns.instructions.md) - Command pattern, pooling, resource management
- 🧠 [MCP Server Guide](instructions/mcp-server-guide.instructions.md) - MCP tools, protocol, error handling
- 🔄 [Development Workflow](instructions/development-workflow.instructions.md) - PR process, CI/CD, security, versioning

---

## What is ExcelMcp?

**ExcelMcp** is a Windows-only toolset for programmatic Excel automation via COM interop, designed for coding agents and automation scripts.

**Four Layers:**
1. **ComInterop** (`src/ExcelMcp.ComInterop`) - Reusable COM automation patterns (STA threading, session management, batch operations, OLE message filter)
2. **Core** (`src/ExcelMcp.Core`) - Excel-specific business logic (Power Query, VBA, worksheets, parameters)
3. **CLI** (`src/ExcelMcp.CLI`) - Command-line interface for scripting
4. **MCP Server** (`src/ExcelMcp.McpServer`) - Model Context Protocol for AI assistants

**Key Capabilities:**
- Power Query M code management (import, export, update, refresh)
- VBA macro management (list, import, export, run)
- Worksheet operations (read, write, create, delete)
- Named range parameters (get, set, create)
- Cell operations (values, formulas)
- Excel instance pooling for MCP server performance

---

## 🎯 Development Quick Start

### Before You Start
1. Read [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) - 5 mandatory rules
2. Check [Testing Strategy](instructions/testing-strategy.instructions.md) for test execution patterns

### Common Tasks
- **Add new command** → Follow patterns in [Architecture Patterns](instructions/architecture-patterns.instructions.md)
- **Excel COM work** → Reference [Excel COM Interop](instructions/excel-com-interop.instructions.md)
- **Modify pool code** → MUST run OnDemand tests (see [CRITICAL-RULES.md](instructions/critical-rules.instructions.md))
- **Add MCP tool** → Follow [MCP Server Guide](instructions/mcp-server-guide.instructions.md)
- **Create PR** → Follow [Development Workflow](instructions/development-workflow.instructions.md)
- **Migrate tests to batch API** → See BATCH-API-MIGRATION-PLAN.md for comprehensive guide
- **Create simple tests** → Use ConnectionCommandsSimpleTests.cs or SetupCommandsSimpleTests.cs as template

### Test Execution
```bash
# Development (fast feedback)
dotnet test --filter "Category=Unit&RunType!=OnDemand"

# Pre-commit (requires Excel)
dotnet test --filter "(Category=Unit|Category=Integration)&RunType!=OnDemand"

# Pool cleanup (MANDATORY when modifying pool code)
dotnet test --filter "RunType=OnDemand"
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
- [Architecture Overview](../docs/ARCHITECTURE-REFACTORING.md)
- [Installation Guide](../docs/INSTALLATION.md)
- [Security Improvements](../docs/SECURITY-IMPROVEMENTS.md)

---

## 🔄 Continuous Learning

After completing significant tasks, update these instructions with lessons learned. See [CRITICAL-RULES.md](instructions/critical-rules.instructions.md) Rule 4.

**Lesson Learned (2025-10-27 - COM Interop Extraction):** Separating COM Interop into standalone project:
1. **New Project Structure:** Created `ExcelMcp.ComInterop` as separate reusable library
2. **Files Moved (Phase 1):** `ComUtilities.cs`, `IOleMessageFilter.cs`, `OleMessageFilter.cs`
3. **Files Moved (Phase 2):** `ExcelSession.cs`, `ExcelBatch.cs`, `ExcelContext.cs`, `ExcelStaExecutor.cs`, `IExcelBatch.cs` (all from Session/)
4. **Tests Moved:** `StaThreadingTests.cs` from `Core.Tests/Unit/Session/` to `ComInterop.Tests/Unit/Session/`
5. **Namespace Changes:** 
   - `Sbroenne.ExcelMcp.Core.ComInterop` → `Sbroenne.ExcelMcp.ComInterop`
   - `Sbroenne.ExcelMcp.Core.Session` → `Sbroenne.ExcelMcp.ComInterop.Session`
   - Test namespace: `Sbroenne.ExcelMcp.Core.Tests.Unit.Session` → `Sbroenne.ExcelMcp.ComInterop.Tests.Unit.Session`
6. **Test Trait Updates:** Changed `[Trait("Layer", "Core")]` to `[Trait("Layer", "ComInterop")]` in StaThreadingTests
7. **Visibility:** Changed `OleMessageFilter` from `internal` to `public` for cross-project use
8. **Bulk Updates:** Used PowerShell for namespace replacements across 40+ files efficiently
9. **Benefits:** ComInterop now provides complete Excel COM automation patterns (utilities, STA threading, session management, batch operations) with its own test suite - other projects can use or exclude entire library
10. **Testing Side Effects:** Tests with Excel process side effects (like `StaThreadingTests`) must use `[Trait("RunType", "OnDemand")]` to avoid running during normal test runs
11. **Session Classes Are Generic:** ExcelSession, ExcelBatch, ExcelStaExecutor are reusable COM interop patterns, not Excel-specific business logic

**Lesson Learned (2025-10-27 - Batch API Migration):** When migrating large test suites to new API patterns:
1. **Strategy Pivot:** Don't force conversion of complex old tests - create NEW simple tests instead
2. **Exclude & Build:** Temporarily exclude unconverted files in .csproj to get clean build fast
3. **Simple Tests Pattern:** Create minimal 1-3 test files per command type that prove API works
4. **CLI Exception Handling:** ALL CLI commands using `BeginBatchAsync` need try-catch wrapping
5. **Missing Using Directives:** Add `using Sbroenne.ExcelMcp.Core.Models;` when using result types
6. **Conversion Helpers:** Convert helpers FIRST before tests that depend on them
7. **Plan Documentation:** Create detailed migration plans for future continuation (see BATCH-API-MIGRATION-PLAN.md)
8. **Test Incrementally:** After each file/group, build and run tests to catch issues early

**Lesson Learned (2025-10-24 - Bulk Refactoring):** When performing bulk refactoring with many find/replace operations:
1. **Preferred:** Use `replace_string_in_file` tool for targeted, unambiguous edits with context
2. **Batch Operations:** Use `grep_search` to find patterns, then use `replace_string_in_file` in parallel for independent changes
3. **Avoid:** PowerShell scripts or terminal commands for code changes - they lack precision and are prone to encoding/parsing issues
4. For large-scale refactorings (100+ replacements), break into smaller batches and test incrementally

**Available Internal Tools (2025-10-24):**
- `replace_string_in_file` - Precise code edits with 3-5 lines of context (use for all code changes)
- `create_file` - Create new files with content (use instead of terminal file creation)
- `read_file` - Read specific line ranges (always check current state before editing)
- `grep_search` - Find patterns across workspace (use to locate code to change)
- `semantic_search` - Find relevant code by intent (use for discovering related code)
- `file_search` - Find files by glob pattern (use to locate files by name/extension)
- `list_dir` - List directory contents (use instead of terminal `ls` or `dir`)
- `get_errors` - Get compile/lint errors (use instead of terminal `dotnet build` for error checking)
- `run_in_terminal` - Execute commands (ONLY for operations with no alternative: dotnet build, dotnet test, git commands)

**Tool Selection Priority:**
1. Code changes → `replace_string_in_file` (always)
2. File creation → `create_file` (always)
3. Find code → `grep_search` or `semantic_search` (always)
4. Check errors → `get_errors` (preferred over terminal build)
5. Build/test/git → `run_in_terminal` (only when no alternative)

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

