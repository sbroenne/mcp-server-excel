# ExcelMcp Project Constitution

> **Governance principles and mandatory rules for all ExcelMcp development**

## 🎯 Project Mission

**ExcelMcp** is a Windows-only toolset for programmatic Excel automation via COM interop, designed for coding agents and automation scripts.

**Core Values:**
- **Reliability**: Zero COM leaks, guaranteed cleanup, tested before commit
- **Developer Experience**: Batch API for performance, clear error messages, comprehensive docs
- **AI-First Design**: MCP Server optimized for LLM consumption
- **Quality Gates**: Zero warnings, all tests pass, comprehensive bug fixes

## ⚠️ CRITICAL RULES (Non-Negotiable)

### Rule 0: NEVER Commit Without Running Tests (CRITICAL)
**NEVER commit, push, or create PRs without first running tests for the code you changed.**

Process:
1. Make code changes
2. Build: `dotnet build` (must succeed with 0 warnings)
3. Run tests: `dotnet test --filter "Feature=<feature>&RunType!=OnDemand"`
4. Verify all tests pass
5. Run pre-commit checks
6. THEN commit

### Rule 1: Success Flag Must Match Reality (CRITICAL)
**NEVER set `Success = true` when `ErrorMessage` is set.**

```csharp
// ❌ CRITICAL BUG
result.Success = true;
result.ErrorMessage = "Query imported but failed...";

// ✅ CORRECT
var result = new OperationResult();
try {
    // ... do work ...
    result.Success = true;  // Only set true on actual success
    return result;
} catch (Exception ex) {
    result.Success = false;  // ✅ Always false in catch!
    result.ErrorMessage = $"Error: {ex.Message}";
    return result;
}
```

**Invariant:** `Success == true` ⟹ `ErrorMessage == null || ErrorMessage == ""`

### Rule 21: Never Commit Automatically (CRITICAL)
**NEVER commit or push code automatically. All commits require explicit user approval.**

## 📐 Architecture Principles

### Four-Layer Design
1. **ComInterop** - Reusable COM patterns (STA threading, sessions, batch, OLE filter)
2. **Core** - Excel business logic (Power Query, VBA, worksheets, parameters)
3. **CLI** - Command-line interface for scripting
4. **MCP Server** - Model Context Protocol for AI assistants

### Command Pattern
- One interface per feature area (IPowerQueryCommands, ISheetCommands, etc.)
- One implementation class per interface
- Partial classes for large implementations (>15 methods)
- All commands accept `IExcelBatch` as first parameter

### Batch API Pattern
**ALL operations use IExcelBatch for exclusive workbook access:**

```csharp
await using var batch = await ExcelSession.BeginBatchAsync(filePath);
await _commands.OperationAsync(batch, args);
await batch.SaveAsync();  // Explicit save required
```

**Benefits:** 75-90% faster than individual operations, guaranteed COM cleanup

### Resource Management
- **WithExcel() Deprecated** - Use Batch API instead
- **COM Cleanup** - try/finally with `ComUtilities.Release(ref obj!)`
- **OLE Message Filter** - Handles cross-thread COM calls
- **Exclusive Access** - One batch per file at a time

## 🧪 Testing Standards

### Test Pyramid
**No Unit Tests** - ExcelMcp has no traditional unit tests (see `docs/ADR-001-NO-UNIT-TESTS.md`)
- **Integration Tests** - Excel COM required, ARE our unit tests
- **OnDemand Tests** - Session/batch infrastructure (slow ~20s each)
- **Manual Tests** - VBA operations (require trust configuration)

### Test Naming
**Pattern**: `MethodName_StateUnderTest_ExpectedBehavior`
- ✅ `List_EmptyWorkbook_ReturnsEmptyList`
- ❌ `List_WithValidFile_ReturnsSuccessResult` (too generic)

See `docs/TEST-NAMING-STANDARD.md` for complete guide.

### File Isolation
- ✅ Each test creates unique file via `CoreTestHelper.CreateUniqueTestFileAsync()`
- ❌ **NEVER** share test files between tests
- ✅ Use `.xlsm` for VBA tests, `.xlsx` otherwise

### Required Traits
```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "PowerQuery")]  // Valid: PowerQuery, DataModel, Tables, etc.
[Trait("RequiresExcel", "true")]
[Trait("RunType", "OnDemand")]  // For slow session/diagnostic tests only
```

### Assertions
- ✅ Binary: `Assert.True(result.Success, $"Reason: {result.ErrorMessage}")`
- ❌ **NEVER** "accept both" patterns
- ✅ **ALWAYS verify actual Excel state** after create/update operations

### SaveAsync Rules (CRITICAL)
- ❌ **FORBIDDEN** unless explicitly testing persistence
- ✅ **ONLY** for round-trip tests: Create → Save → Re-open → Verify
- ❌ **NEVER** call in middle of test (breaks subsequent operations)

## 📝 Documentation Standards

### Documentation Hierarchy
- **Root Level** - Essential user-facing only (README, LICENSE, SECURITY)
- **docs/** - Permanent implementation/process documentation
- **specs/** - Feature specifications (Spec Kit format)
- **src/Component/** - Component-specific docs

### Spec Kit Structure
```
specs/###-feature-name/
├── spec.md   # WHAT to build (user stories, requirements)
└── plan.md   # HOW it's built (architecture, decisions)
```

### Naming Conventions
- ✅ `TOPIC-NAME.md` (ALL CAPS for discoverability)
- ✅ `ADR-NNN-DECISION-NAME.md` (Architecture Decision Records)
- ❌ NO temporary files at root (SUMMARY.md, FIX.md, TESTS.md)

## 🔐 Security Standards

### COM Security
- **No Credential Storage** - Connections use existing Excel sources
- **DAX Injection Risk** - Low, TOM API validates syntax
- **File Access** - Validated via `FileAccessValidator` before operations
- **COM Cleanup** - Guaranteed via try/finally

### VBA Security
- **VBA Trust Required** - Programmatic access blocked by default
- **Development Only** - Never enable VBA trust in production
- **VBA Code Validation** - None, execute only trusted macros

### Connection String Sanitization
```csharp
// ALWAYS sanitize before output
string safe = ConnectionHelpers.SanitizeConnectionString(rawConnectionString);
```

## 🚀 Performance Requirements

### Batch API Performance
- Individual operations (10x): ~10-15 seconds
- Batch operations (10x): ~2-3 seconds
- **Target**: 75-85% improvement

### Operation Timeouts
- **Default**: 2 minutes (most operations)
- **Extended**: 5 minutes (refresh, data model, large ranges)
- **Configurable**: Per-operation override via timeout parameter

### Bulk Operations
- Use `Range.Value2` for 2D arrays, not cell-by-cell
- SetNumberFormatsAsync accepts 2D array for cell-by-cell formats
- CreateBulkAsync for named ranges (90% faster than individual)

## 🔄 Git Workflow

### Branch Strategy
- **main** - Protected, requires PR + CI/CD pass
- **feature/** - Feature development
- **fix/** - Bug fixes

### PR Requirements
1. Build passes (0 warnings, 0 errors)
2. All tests pass (relevant feature tests + OnDemand if session code changed)
3. Docs updated (tool/method docs, user docs, prompts)
4. Pre-commit hooks pass (COM leaks, success flag, coverage)
5. **Automated review comments fixed** (Copilot, GitHub Security)

### Commit Standards
- Descriptive messages explaining WHY
- Reference issue numbers
- No TODO/FIXME/HACK markers
- No commented-out code (use git history)

## 🐛 Bug Fix Standards

**Every bug fix MUST include all 6 components:**
1. ✅ **Code Fix** - Minimal surgical changes
2. ✅ **Tests** - Minimum 5-8 tests (regression + edge + backwards compat + MCP)
3. ✅ **Documentation** - Update 3+ files (tool docs, user docs, prompts)
4. ✅ **Workflow Hints** - Update SuggestedNextActions, error messages
5. ✅ **Quality Verification** - Build passes, tests green, 0 warnings
6. ✅ **PR Description** - Comprehensive summary (bug, fix, tests, docs)

See `.github/instructions/bug-fixing-checklist.instructions.md` for complete process.

## 🛠️ Code Standards

### .NET Class Design
- **One Public Class Per File** - Standard .NET practice
- **File Name = Class Name** - `RangeCommands.cs` contains `RangeCommands`
- **Partial Classes** - Split 15+ method classes by feature domain
- **Folder = Organization** - `Commands/Range/RangeCommands.cs`

### Naming Conventions
- **Descriptive Names** - `RangeCommands` ✅, `Commands` ❌
- **Async Methods** - End with `Async` suffix
- **Private Fields** - `_camelCase` with underscore prefix
- **Constants** - `PascalCase` for public, `UPPER_SNAKE_CASE` for internal

### Error Handling
- **MCP Tools** - Return JSON with `success: false` for business errors
- **Throw McpException** - Only for parameter validation and pre-conditions
- **Core Commands** - Return `OperationResult` with `Success` flag
- **Enrich TimeoutException** - Add LLM guidance fields

### Excel COM Patterns
- **Use Late Binding** - `dynamic` types
- **1-Based Indexing** - Excel collections start at 1
- **Release COM Objects** - `ComUtilities.Release(ref obj!)`
- **QueryTable Refresh REQUIRED** - `.Refresh(false)` synchronous for persistence
- **NEVER RefreshAll()** - Async/unreliable, use individual refresh

## 📊 Quality Gates

### Build Requirements
- **Zero Warnings** - `TreatWarningsAsErrors=true`
- **Analyzers Enabled** - Security rules as errors
- **StyleCop** - Enforced code style

### Pre-Commit Checks
- `check-com-leaks.ps1` - Must report 0 leaks
- `check-success-flag.ps1` - Detect Success/ErrorMessage violations
- No TODO/FIXME/HACK markers
- No commented-out code

### CI/CD Gates
- All builds pass (MCP Server, CLI, Core, Tests)
- All integration tests pass (Azure self-hosted runner)
- CodeQL security scan passes
- Dependency review passes

## 🔧 Development Workflow

### Test Execution
```bash
# Fast feedback (excludes VBA, excludes OnDemand)
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Specific feature
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"

# Session/batch changes (MANDATORY)
dotnet test tests/ExcelMcp.ComInterop.Tests/ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"

# VBA tests (manual only, requires VBA trust)
dotnet test --filter "(Feature=VBA|Feature=VBATrust)&RunType!=OnDemand"
```

### Pre-Commit Workflow
1. Search TODO/FIXME/HACK
2. Delete commented-out code
3. Run relevant tests
4. Check docs updated
5. Run pre-commit hooks: `.\scripts\pre-commit.ps1`

### PR Review Workflow
1. Create PR
2. **Immediately check for automated review comments** (within 1-2 min)
3. Fix all Copilot and GitHub Security comments
4. Push fixes to PR branch
5. Request human review only after automated issues resolved

## 📚 Documentation Practices

### README Management
- **Three READMEs** - Root (comprehensive), McpServer (NuGet), VSCode (extension)
- **Tool Counts Must Match** - Verify against code
- **Safety Messaging** - COM API benefits, zero corruption
- **Version Numbers** - Auto-managed by release workflow (NEVER manual)

### MCP LLM Guidance
- **Tool Descriptions** - Server-specific behavior, non-enum parameters
- **Prompt Files** - Detailed workflows, action disambiguation
- **Completions** - Freeform string parameters only (NOT enums)
- **Elicitations** - Pre-flight checklists for complex operations

### Code Comments
- **XML Comments** - All public APIs
- **Implementation Comments** - Explain WHY, not WHAT
- **TODOs** - Forbidden before commit
- **References** - Link to Microsoft docs, specs, ADRs

## 🎓 Key Lessons Learned

### Success Flag Invariant
- Success=true means NO errors (not even warnings)
- Set Success in try block, always false in catch
- Pre-commit hook detects violations

### Batch API Evolution
- Replace WithExcel() with BeginBatchAsync
- Explicit save via SaveAsync
- Create NEW simple tests (not update old ones)

### Excel Quirks
- Type 3/4 both handle TEXT connections
- RefreshAll() unreliable, use individual refresh
- Numeric properties return double, not int/enum
- QueryTable.Refresh(false) required for persistence

### MCP Design Philosophy
- Prompts are shortcuts, not tutorials
- LLMs know Excel/programming
- Focus on server-specific behavior
- Tool descriptions for always-visible guidance

### Testing Strategy
- Integration tests ARE unit tests (no mocking)
- Each test gets unique file
- Binary assertions, no "accept both"
- SaveAsync only for persistence tests
- Round-trip validation for CRUD

## 🔍 References

### Core Documentation
- [CRITICAL-RULES.md](.github/instructions/critical-rules.instructions.md) - 21 mandatory rules
- [Testing Strategy](.github/instructions/testing-strategy.instructions.md) - Test patterns
- [Excel COM Interop](.github/instructions/excel-com-interop.instructions.md) - COM patterns
- [MCP Server Guide](.github/instructions/mcp-server-guide.instructions.md) - MCP implementation
- [Bug Fixing Checklist](.github/instructions/bug-fixing-checklist.instructions.md) - 6-step process

### Architecture Decisions
- [ADR-001: No Unit Tests](docs/ADR-001-NO-UNIT-TESTS.md)
- [Architecture Patterns](.github/instructions/architecture-patterns.instructions.md)
- [Timeout Implementation Guide](docs/TIMEOUT-IMPLEMENTATION-GUIDE.md)

### Process Documentation
- [Development Workflow](.github/instructions/development-workflow.instructions.md)
- [Release Strategy](docs/RELEASE-STRATEGY.md)
- [Pre-Commit Setup](docs/PRE-COMMIT-SETUP.md)

---

**Last Updated**: 2025-01-10

