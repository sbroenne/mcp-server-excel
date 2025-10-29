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
- **Range Operations** (Phase 1 implementation in progress) - Unified API for all range data operations (get/set values/formulas, clear variants, find/replace, sort, insert/delete, copy/paste, UsedRange, CurrentRegion, hyperlinks)
- Power Query M code management (import, export, update, refresh)
- VBA macro management (list, import, export, run)
- Worksheet lifecycle management (list, create, rename, copy, delete)
- Named range parameters (create, delete, update, list, get/set single values)
- Data Model operations (list tables/measures/relationships, export measures, refresh, delete)
- Connection management (list, view, import/export, update, refresh, test, properties)

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
- **Range API implementation** → See [Range API Specification](../specs/RANGE-API-SPECIFICATION.md) for complete design (38 methods, MCP-first, breaking changes acceptable)

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
- [Installation Guide](../docs/INSTALLATION.md)
- [Range API Specification](../specs/RANGE-API-SPECIFICATION.md) - Comprehensive design for unified range operations (Phase 1 implementation)
- [Range Refactoring Analysis](../specs/RANGE-REFACTORING-ANALYSIS.md) - LLM perspective on consolidating fragmented commands

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

**Lesson Learned (2025-10-27 - Excel Type 3/4 Confusion):** When Excel COM API reports unexpected types:
1. **Root Cause Investigation:** Excel returns type 4 (WEB) for TEXT connections created with "TEXT;filepath"
2. **This is NOT a bug:** Excel COM API behavior, not a code defect - must be accepted
3. **Dual Handling Pattern:** Handle BOTH type 3 AND type 4 in ALL connection property methods
4. **Try/Catch Fallback:** Try TextConnection first, fall back to WebConnection if that fails
5. **Test Reality, Not Ideals:** Update test expectations to match Excel's actual behavior (expect "WEB" not "TEXT")
6. **Graceful Degradation:** Allow operations to succeed even if some properties aren't settable (Excel limitation)
7. **Pragmatic Solutions:** Accept quirky behavior instead of fighting it - results in cleaner code
8. **Comprehensive Updates:** When fixing type handling, update ALL related methods consistently (6 methods updated)
9. **Pattern Consistency:** Use same try/catch pattern across all property access methods for maintainability

**Lesson Learned (2025-10-27 - MCP Prompt Design):** When creating prompts for MCP servers:
1. **Research First:** Study real MCP servers (fetch, everything, time) to understand best practices
2. **Prompts ≠ Tutorials:** MCP prompts should be SHORT user shortcuts, not 400+ line programming tutorials
3. **LLMs Know Programming:** Don't teach TypeScript, M code, or VBA - LLMs already know these languages
4. **Domain Knowledge Only:** Keep prompts focused on domain-specific facts LLMs can't infer (Excel connection types, COM API limitations)
5. **Quality > Quantity:** Deleted 4 tutorial prompts (1,538 lines), kept 1 reference (54 lines) - better results
6. **Prompt Purpose:** Help users invoke tools efficiently, not educate LLMs on general programming
7. **Pattern Knowledge:** Don't document patterns LLMs understand (batching, transactions) - use tool descriptions instead
8. **Validation:** If it reads like a tutorial or "how to code X", it's wrong for MCP prompts

**Lesson Learned (2025-10-27 - Range API Design):** When designing unified APIs to replace fragmented commands:
1. **Specification First:** Create comprehensive specification BEFORE implementation - iterate with validations from multiple perspectives (power user, AI agent, LLM usability)
2. **LLM Perspective Analysis:** "As an LLM using the MCP server" reveals UX issues - separate methods for named ranges vs explicit addresses = unnecessary complexity
3. **Single Cell = Range Principle:** Don't create separate cell API - single cell is just 1x1 range (consistent data format: always 2D arrays)
4. **Separation of Concerns:** Lifecycle operations (create, delete, rename worksheet) ≠ Data operations (read, write, clear range) - split into separate tools/commands
5. **Named Range Transparency:** Excel COM resolves named ranges natively (`Worksheet.Range("SalesData")` works like `Worksheet.Range("A1:D10")`) - API should reflect this (one rangeAddress parameter accepts both)
6. **COM-Backed Only:** Don't implement data processing in server (transpose, statistics) - if Excel COM doesn't provide it, don't add it
7. **Breaking Changes Strategy:** Clean architecture > backwards compatibility - delete fragmented commands entirely (still in minor releases, breaking changes acceptable during active development)
8. **MCP-First Implementation:** Implement MCP server before CLI - faster feedback loop, JSON simpler than CSV conversion
9. **Comprehensive Refactoring Analysis:** Document what gets deleted, what gets refactored, what gets added - file-by-file impact analysis prevents surprises
10. **Migration Examples:** Provide before/after examples for every breaking change - LLMs and users need clear migration path

**Lesson Learned (2025-10-27 - Range API Refactoring Scope):** Analyzing existing commands for consolidation:
1. **SheetCommands Data Operations:** `ReadAsync`, `WriteAsync`, `ClearAsync`, `AppendAsync` are range operations disguised as sheet operations - move to RangeCommands
2. **CellCommands Redundancy:** All 4 methods (`GetValue`, `SetValue`, `GetFormula`, `SetFormula`) are 1x1 range operations - delete entire interface
3. **Unified Interface Benefits:** From LLM perspective, one tool (`excel_range`) for all data operations eliminates "which tool to use?" confusion
4. **Lifecycle vs Data Split:** SheetCommands reduced to 5 lifecycle-only methods (List, Create, Rename, Copy, Delete) - clearer responsibilities
5. **Power User Operations:** Excel COM provides Find, Sort, Insert/Delete rows, UsedRange, CurrentRegion - expose all of them, don't cherry-pick
6. **38 Methods Is Acceptable:** Large interface is fine if it's cohesive and well-organized (values, formulas, clear, copy, insert/delete, find/sort, discovery, hyperlinks)
7. **CSV Conversion Layer:** Keep CSV handling in CLI layer only - Core uses `List<List<object?>>`, MCP uses JSON arrays, CLI converts CSV ↔ 2D arrays
8. **Testing Deletion Strategy:** Don't try to migrate complex old tests - delete them and create NEW simple tests that prove unified API works
9. **Impact Documentation:** 13 actions deleted (worksheet.read/write/clear/append + all cell actions), 38 actions added (all range operations) - net improvement in capabilities
10. **Acceptable Timeline:** ~1-2 weeks for Phase 1 (Core + MCP + CLI minimal + CLI full) is reasonable for 38-method implementation

**Lesson Learned (2025-10-27 - Range API Implementation - File Organization & Refactoring Strategy):** When implementing large APIs with breaking changes:

**1. File Organization - Context-Specific Patterns:**
- **Core Business Logic** (Commands): Use partial classes when files exceed ~500 lines
  - Example: RangeCommands split into 9 partial files (Values, Formulas, Clear, Copy, Editing, Search, Discovery, Hyperlinks, Named Ranges)
  - Benefit: ~200 lines per file, git-friendly, team-friendly, feature-focused
- **MCP Translation Layers** (Tools): Single file acceptable up to ~1400 lines
  - Example: ExcelRangeTool.cs (30 actions, ~1400 lines) remains single file
  - Rationale: LLM needs discoverability > organization; action-based routing simple to navigate
  - Pattern: Follows ExcelWorksheetTool, ExcelPowerQueryTool precedent
- **Key Insight**: Translation layers ≠ business logic; different optimization goals

**2. Large-Scale Refactoring Strategies:**
- **File Recreation > Incremental Edits** when removing 50%+ of content
  - Example: SheetCommands.cs (430→186 lines) recreated from scratch
  - Example: ExcelWorksheetTool.cs (428→220 lines) recreated cleanly
  - Benefit: Cleaner result than complex `replace_string_in_file` operations
  - Issue: Multi-line text replacements prone to whitespace/formatting mismatches
- **When to Use `replace_string_in_file`:**
  - Small targeted changes (single method, property rename)
  - Requires exact whitespace matching (3-5 lines of context)
  - Best for surgical edits, not structural refactoring
- **File Recreation Process:**
  1. Read original file to understand structure
  2. Create new file with only desired functionality
  3. Preserve XML documentation, using directives, namespace
  4. Build and fix compilation errors incrementally

**3. Breaking Change Documentation:**
- **Commit Message Structure** (from Phase 1A commit):
  - Clear title stating scope: "Phase 1A: Refactor SheetCommands to lifecycle-only"
  - BREAKING CHANGES section listing all API removals
  - Before/after metrics (line counts, method counts, action counts)
  - Migration paths documented (old API → new API)
  - Note expected compilation errors during phased refactoring
- **Impact Transparency:**
  - Document total lines removed (~650 lines across 6 files in Phase 1A)
  - List affected components (Core, MCP Server, Tests)
  - Specify which errors are intentional (CLI Phase 1B expected)

**4. Test Dependency Management:**
- **Cross-Feature Dependencies** require careful tracking
  - Example: PowerQueryWorkflowGuidanceTests used SheetCommands.ReadAsync
  - Solution: Update to RangeCommands.GetValuesAsync when SheetCommands refactored
  - Pattern: Search for deleted method names before removing
- **Test Refactoring Strategy:**
  - Delete tests for removed functionality (don't migrate)
  - Create NEW tests for unified API (fresh perspective)
  - Example: Deleted 4 SheetCommands data tests, Range API has 24 comprehensive tests
- **Model Changes Propagate:**
  - Changed WorksheetDataResult → RangeValueResult
  - Updated property access: `.Data` → `.Values`
  - Required updates across test files

**5. Phased Refactoring Pattern:**
- **Phase 1A** (Core + MCP Server):
  - Implement new unified API (RangeCommands, ExcelRangeTool)
  - Delete obsolete commands (CellCommands, HyperlinkCommands)
  - Refactor overlapping commands (SheetCommands lifecycle-only)
  - Update MCP Server tool inventory
  - **Accept CLI compilation errors** - documented as expected
- **Phase 1B** (CLI Implementation):
  - Create CLI wrapper for new API
  - Fix CLI compilation errors from Phase 1A
  - Update routing and documentation
  - Complete user-facing integration
- **Benefit**: Faster Core/MCP iteration without CLI complexity

**6. MCP Tool Design Principles:**
- **Single File Acceptable** for translation layers
  - LLMs prioritize discoverability over file organization
  - Action-based routing (switch statement) easy to navigate
  - ~1400 lines manageable for AI comprehension
- **Action Density**: 30 actions ≈ ~47 lines per action (acceptable)
- **Validation**: Regex pattern at top documents available actions
- **Pattern Consistency**: Follow existing tools (ExcelWorksheetTool pattern)

**7. Architecture Documentation Importance:**
- **Critical**: Update architecture-patterns.instructions.md with official Microsoft guidelines
  - Link to Microsoft Learn documentation
  - Clarify when to use partial classes vs single files
  - Document patterns with code examples
- **Prevents Confusion**: AI needs explicit guidance on file organization
  - "One class per file" is .NET standard
  - Partial classes documented exception for large classes
  - Context matters: business logic ≠ translation layer

**8. Metrics & Validation:**
- **Success Criteria** from Phase 1A:
  - Core: 38 methods implemented (9 partial files)
  - MCP: 30 actions implemented (1 file)
  - Tests: 24/24 passing (100% success rate)
  - Build: 0 errors, 0 warnings (Core + MCP)
  - Code reduction: ~650 lines removed
  - Compilation errors: 4 expected in CLI (Phase 1B)
- **Timeline**: Phase 1A completed in ~1 week (spec → implementation → testing → refactoring → deletions → MCP tool → SheetCommands refactoring)

**Lesson Learned (2025-10-28 - Phase 1B CLI Implementation & Testing Strategy):** When implementing CLI wrappers for Core commands:
1. **Don't Duplicate Integration Tests**: If Core integration tests cover business logic with Excel COM, CLI doesn't need to re-test the same operations
2. **CLI Testing Focus**: Only test CLI-specific concerns (argument parsing, exit codes, CSV conversion helpers)
3. **Manual Testing Suffices**: Quick manual verification (CSV round-trip, formula operations, single cells) proves CLI wrapper works
4. **Core Tests Are Authoritative**: 24 passing Core integration tests = business logic verified; CLI just formats I/O
5. **CSV Conversion Pattern**: CLI layer converts CSV ↔ 2D arrays (`List<List<object?>>`); Core uses native 2D arrays; MCP uses JSON
6. **Type Inference**: ParseCsvTo2DArray auto-detects numbers, booleans, nulls (empty strings) vs strings for better UX
7. **Progressive Implementation**: Start with 7 essential commands (get/set values/formulas, clear variants), add 23 more later if needed
8. **Migration Documentation Critical**: Users need clear old→new command mapping (COMMANDS.md migration guide essential)
9. **Help Text Updates**: Update CLI help text synchronously with command additions (examples section + command list)
10. **Commit Message Simplicity**: Long multi-paragraph commit messages fail; use short title + 1-2 sentence body instead
11. **Scope Discipline**: Stick to defined phase boundaries - don't add "nice to have" features beyond the plan, even if Core/MCP already support them

**Lesson Learned (2025-10-28 - Scope Management & Phase Boundaries):** When implementing multi-phase features:
1. **Define clear phase boundaries** - Phase 1B = "7 essential commands to replace deleted functionality"
2. **Progressive implementation is intentional** - Start with minimum viable set, add more based on user demand
3. **Resist scope creep** - Don't add features just because they're "easy" or already implemented in other layers
4. **Each phase should solve a specific problem** - Phase 1B solved: "Users lost sheet-read/write/clear commands"
5. **Future additions go in separate PRs** - Remaining 23 CLI commands are user-driven, not automatic
6. **Complete means complete** - When phase objectives are met, mark it done and move on

**Lesson Learned (2025-01-29 - Spec Validation Against Official Documentation):** When encountering architectural decisions based on "API limitation" claims:

**Situation**: DataModelCommands refactoring - Original spec claimed: "Excel COM API limited, use TOM for CREATE/UPDATE operations"

**User Instinct**: "I do not trust our spec" → Request validation against Microsoft official documentation

**Agent Research Process**:
1. **Microsoft Docs Search**: "Excel Data Model Power Pivot DAX measures programmatic access COM API"
2. **Microsoft Docs Search**: "Excel VBA Model object ModelMeasures ModelTables ModelRelationships COM automation"
3. **Fetch API Documentation**: https://learn.microsoft.com/en-us/office/vba/api/excel.modelmeasures.add
4. **Fetch API Documentation**: https://learn.microsoft.com/en-us/office/vba/api/excel.modelrelationships.add
5. **Microsoft Docs Search**: "Tabular Object Model TOM Analysis Services embedded Excel workbook"

**Critical Discovery**: Original spec was **COMPLETELY WRONG**

**Actual Truth** (Microsoft Official Documentation):
- ✅ Excel COM API **FULLY SUPPORTS** `ModelMeasures.Add()` since Office 2016
- ✅ Excel COM API **FULLY SUPPORTS** `ModelRelationships.Add()` since Office 2016
- ✅ Excel COM API **FULLY SUPPORTS** updating measures (Formula, Description, FormatInformation properties)
- ✅ Excel COM API **FULLY SUPPORTS** updating relationships (Active property)
- ❌ TOM **ONLY REQUIRED** for advanced features (calculated columns, hierarchies, perspectives, KPIs)

**Impact**:
- Saved weeks of unnecessary TOM integration work (NuGet packages, connection management, deployment scenarios)
- Simpler implementation (native Excel operations, no server dependencies)
- Better user experience (works offline, no external services)
- Complete CRUD capability using Excel COM only

**Mandatory Process for Future Work**:
1. ✅ **Always search Microsoft official documentation FIRST** (use mcp_microsoft_doc tools)
2. ✅ **Fetch specific API reference pages** (validate exact method signatures and parameters)
3. ✅ **Never trust secondary sources or assumptions** (even internal specs can be wrong)
4. ✅ **Update specs immediately when errors found** (correct misinformation)
5. ✅ **Archive incorrect specs** (preserve history with .archived.md suffix, prevent future confusion)
6. ✅ **Document lessons learned** (add to copilot instructions for future AI sessions)

**Examples of Validated APIs**:
```csharp
// Microsoft Official: ModelMeasures.Add() method
dynamic measures = table.ModelMeasures;
dynamic newMeasure = measures.Add(
    MeasureName: "TotalSales",
    AssociatedTable: table,
    Formula: "SUM(Sales[Amount])",
    FormatInformation: model.ModelFormatCurrency,
    Description: "Total sales amount"
);

// Microsoft Official: ModelRelationships.Add() method
dynamic relationships = model.ModelRelationships;
relationships.Add(
    ForeignKeyColumn: salesTable.ModelTableColumns.Item("CustomerID"),
    PrimaryKeyColumn: customersTable.ModelTableColumns.Item("ID")
);

// Microsoft Official: Property updates
measure.Formula = "CALCULATE(SUM(Sales[Amount]))";  // Read/Write
measure.Description = "Updated description";         // Read/Write
relationship.Active = false;                         // Read/Write
```

**Red Flags That Require Validation**:
- Claims like "API X doesn't support operation Y" (verify with official docs)
- "Use library Z because native API insufficient" (validate limitation exists)
- "Requires external service/package for basic operation" (check if native alternative exists)
- Any architectural decision based on undocumented assumptions

**Principle**: **Microsoft official documentation is ALWAYS authoritative over specs, blog posts, Stack Overflow, or assumptions**

**Lesson Learned (2025-10-29 - QueryTable Persistence Bug):** When debugging COM interop issues with mysterious persistence failures:
1. **Symptoms Recognition**: If objects exist in memory but disappear after file reopen, suspect async operations
2. **Research Pattern**: Search Microsoft official docs for VBA examples showing proven patterns
3. **RefreshAll() Caveat**: `RefreshAll()` is ASYNCHRONOUS for objects with `BackgroundQuery=true`
4. **Individual Refresh Required**: QueryTables must call `.Refresh(false)` synchronously to persist properly
5. **Microsoft VBA Examples Are Gold**: Official VBA code samples show production-proven patterns (Create → Refresh(False) → Save)
6. **Debug At Core Level**: Create simple Core-level tests to isolate COM behavior from MCP/CLI layers
7. **Async vs Sync Matters**: Excel COM has both sync and async variants - wrong choice causes silent failures
8. **Save Isn't Enough**: Some objects need explicit initialization (like Refresh) before Save to persist
9. **Document Discovery**: Add critical findings to excel-com-interop.instructions.md immediately
10. **Trust User Instincts**: When user says "research this online", they're often sensing a pattern mismatch

**Key Insight**: RefreshAll() claims to refresh queries but doesn't properly initialize individual QueryTables for disk persistence. Individual queryTable.Refresh(false) is mandatory.

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

**Pre-Commit Checklist:**
1. ✅ Search for TODO/FIXME/HACK markers: `grep_search` with pattern `//\s*(TODO|FIXME|HACK|XXX)`
2. ✅ Resolve ALL markers before committing (see CRITICAL-RULES.md Rule 7)
3. ✅ Delete commented-out code (use git history if needed)
4. ✅ Verify all tests pass
5. ✅ Update documentation if behavior changed

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

