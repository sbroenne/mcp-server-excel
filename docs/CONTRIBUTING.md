# Contributing to ExcelMcp

Thank you for your interest in contributing to Sbroenne.ExcelMcp! This project is designed to be extended by the community, especially to support coding agents like GitHub Copilot.

## 🎯 Project Vision

ExcelMcp aims to be the go-to command-line tool for coding agents to interact with Microsoft Excel files. We prioritize:

- **Simplicity** - Clear, predictable commands
- **Reliability** - Robust COM automation
- **Extensibility** - Easy to add new features
- **Agent-Friendly** - Designed for AI coding assistants

## 🚀 Getting Started

### Development Environment

1. **Prerequisites**:
   - Windows OS (required for Excel COM)
   - Visual Studio 2022 or VS Code
   - .NET 10 SDK
   - Microsoft Excel installed

2. **Setup**:
   ```powershell
   git clone https://github.com/sbroenne/mcp-server-excel.git
   cd mcp-server-excel
   dotnet restore
   dotnet build
   ```

3. **Test your setup** (surgical — don't run the full integration suite, it takes 45+ minutes):
   ```powershell
   dotnet test --filter "Feature=Sheet&RunType!=OnDemand"
   ```

## 🚨 **CRITICAL: Pull Request Workflow Required**

**All changes must be made through Pull Requests (PRs).** Direct commits to `main` are prohibited.

**Merge Strategy: Squash Merge** — All PRs are merged via squash merge (single commit to `main`). This keeps the history clean.

### Quick PR Process

1. **Create feature branch**: `git checkout -b feature/your-feature`
2. **Make changes**: Code, tests, documentation
3. **Run the pre-commit hook**: install it once with `Copy-Item scripts\pre-commit.ps1 .git\hooks\pre-commit`, then let it run on every commit — it enforces 14 automated gates (COM leak detection, MCP/CLI coverage parity, Release build, packaging deliverables, smoke tests, and more). Never bypass it with `--no-verify`.
4. **Push branch**: `git push origin feature/your-feature`
5. **Create PR**: Use GitHub's PR template
6. **Address review**: Make requested changes, including any automated review comments (Copilot, GitHub Advanced Security)
7. **Merge**: After approval and CI checks pass — **GitHub will squash commits automatically**
   - Verify the final commit message accurately describes the changes
   - After merge, your feature branch can be safely deleted

📋 **Detailed workflow**: See [DEVELOPMENT.md](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/DEVELOPMENT.md) for complete instructions.

## 📋 Development Guidelines

### Code Style

- **C# 12** features encouraged (file-scoped namespaces, records, pattern matching)
- **Nullable reference types** enabled - handle nulls properly
- **No warnings** - project must build with zero warnings
- **XML documentation** for public APIs (these docs are extracted into MCP tool descriptions and shown to LLMs — keep them accurate)
- **Consistent naming** - follow established patterns

### Architecture

ExcelMcp has **two equal entry points** — an MCP Server and a CLI — sharing one Core layer:

```
MCP Server ──► In-process ExcelMcpService ──► Core Commands ──► Excel COM
CLI ─────────► CLI Daemon (named pipe) ─────► Core Commands ──► Excel COM
```

- **`ExcelMcp.ComInterop`** - Reusable COM automation primitives (STA threading, session/batch management)
- **`ExcelMcp.Core`** - Excel business logic (Power Query, VBA, worksheets, PivotTables, etc.)
- **`ExcelMcp.Service`** - Excel session management and command routing
- **`ExcelMcp.CLI`** - Command-line interface (session-based: `excelcli session open`, then operate on the session, then `excelcli session close --save`)
- **`ExcelMcp.McpServer`** - Model Context Protocol tools for AI assistants
- **`ExcelMcp.Generators*`** - Source generators that produce CLI commands and MCP tools directly from Core interfaces — you do **not** hand-write CLI verb registration or MCP tool schemas

#### Command Pattern

Core Commands use the batch API and let exceptions propagate — never wrap `batch.Execute()` in a try-catch that returns an error result:

```csharp
public DataType MyOperation(IExcelBatch batch, string arg1)
{
    return batch.Execute((ctx, ct) =>
    {
        dynamic? item = null;
        try
        {
            item = ctx.Book.SomeObject;
            // ... operation logic ...
            return someData;
        }
        finally
        {
            ComUtilities.Release(ref item!); // COM cleanup only — no catch block here
        }
    });
    // batch.Execute() catches exceptions via TaskCompletionSource and
    // returns OperationResult { Success = false, ErrorMessage } automatically
}
```

#### Critical Rules

1. **Always use the batch API** - Never manage Excel lifecycle manually
2. **Excel uses 1-based indexing** - `collection.Item(1)` is the first element
3. **Never suppress exceptions** with a catch block that returns `Success = false` — let `batch.Execute()` handle it
4. **`Success = true` must never coexist with a non-empty `ErrorMessage`**
5. **COM objects** are released only in `finally` blocks, never swallowed in empty `catch` blocks

### Excel COM Best Practices

- **Late binding with dynamic types** for COM interop
- **Proper error handling** - Catch `COMException` where specific handling is needed; otherwise let exceptions propagate
- **Resource cleanup** - Batch API handles COM object lifecycle automatically; release ad-hoc `dynamic` COM objects yourself in `finally`
- **Input validation** - Check file existence and argument validity early

### Testing

ExcelMcp uses **integration tests only** — no unit tests, since COM interop bugs (STA threading, leaks, type conversion) only manifest against a real Excel instance. Follow TDD: write a failing test first, watch it fail, then implement.

```powershell
# Surgical, feature-scoped testing (2-5 minutes) — always prefer this over the full suite
dotnet test --filter "Feature=PowerQuery&RunType!=OnDemand"

# Full non-VBA suite (10-15 minutes) — only when you need broad confidence
dotnet test --filter "Category=Integration&RunType!=OnDemand&Feature!=VBA&Feature!=VBATrust"

# Session/batch changes require the slower OnDemand suite too
dotnet test --filter "RunType=OnDemand"
```

Before submitting a PR:

1. Tests pass for the feature(s) you changed
2. Excel process cleanup verified - no `excel.exe` remains after tests finish
3. Error conditions tested (missing files, invalid arguments, etc.)
4. Build has zero warnings
5. Pre-commit hook passes (all 14 gates)

## 🔧 Adding a New Operation

New operations are added to the **Core** interface/implementation; CLI commands and MCP tool schemas are then generated automatically — you don't hand-write CLI arg parsing or MCP tool registration.

1. **Add the method to the relevant Core interface** (e.g. `Commands/Sheet/ISheetCommands.cs`), with XML doc comments (these become the MCP tool/parameter descriptions).
2. **Implement it** in the corresponding partial class (e.g. `SheetCommands.Lifecycle.cs`), following the batch-API pattern above.
3. **Build the solution** - the source generators (`ExcelMcp.Generators`, `ExcelMcp.Generators.CLI`) produce the CLI verb and MCP tool automatically from the interface.
4. **Add integration tests** for the new operation (TDD: write them first).
5. **Update `FEATURES.md`** with the new operation and its updated operation count — `scripts/check-doc-counts.ps1` enforces that documented counts match the code.

## 📝 Pull Request Process

### Before Submitting

- [ ] Code builds with zero warnings
- [ ] Feature-scoped tests pass (`dotnet test --filter "Feature=<name>&RunType!=OnDemand"`)
- [ ] Excel processes clean up properly
- [ ] Added appropriate error handling (no suppressed exceptions)
- [ ] Updated `FEATURES.md` / relevant docs if the operation count or behavior changed
- [ ] Pre-commit hook passes locally

### PR Description Template

```markdown
## Summary
Brief description of changes

## Type of Change
- [ ] Bug fix
- [ ] New feature
- [ ] Breaking change
- [ ] Documentation update

## Testing
- [ ] Tested manually with Excel files
- [ ] Verified Excel process cleanup
- [ ] Tested error conditions
- [ ] VBA script execution tested (if applicable)
- [ ] No build warnings

## Checklist
- [ ] Code follows project conventions
- [ ] Self-review completed
- [ ] Updated documentation as needed
```

## 🎨 UI Guidelines

### Spectre.Console Usage

```csharp
// Success (green checkmark)
AnsiConsole.MarkupLine($"[green]✓[/] Operation succeeded");

// Error (red)  
AnsiConsole.MarkupLine($"[red]Error:[/] {message.EscapeMarkup()}");

// Warning (yellow)
AnsiConsole.MarkupLine($"[yellow]Note:[/] {message}");

// Info/debug (dim)
AnsiConsole.MarkupLine($"[dim]{message}[/]");

// Headers (cyan)
AnsiConsole.MarkupLine($"[cyan]{title}[/]");
```

### Output Consistency

- **Tables** for structured data (query lists, sheet lists)
- **Panels** for code blocks (M code display)
- **Progress indicators** for long operations
- **Clear error messages** with actionable guidance

## 🐛 Bug Reports

When reporting bugs, please include:

- **Excel version** and Windows version
- **Command used** and arguments
- **Expected behavior** vs actual behavior
- **Sample Excel file** (if possible)
- **Error messages** (full text)

## 💡 Feature Requests

Great feature requests include:

- **Use case description** - Why is this needed?
- **Proposed command syntax** - How should it work?
- **Excel operations involved** - What APIs would be used?
- **Target users** - Coding agents? Direct users?

## 📚 Learning Resources

- [Excel VBA Object Model Reference](https://docs.microsoft.com/en-us/office/vba/api/overview/excel)
- [Power Query M Language Reference](https://docs.microsoft.com/en-us/powerquery-m/)
- [Spectre.Console Documentation](https://spectreconsole.net/)
- [.NET COM Interop Guide](https://docs.microsoft.com/en-us/dotnet/framework/interop/interoperating-with-unmanaged-code)

## 📦 For Maintainers

- [NuGet Publishing Guide](https://github.com/sbroenne/mcp-server-excel/blob/main/docs/NUGET-GUIDE.md) - Complete guide for publishing all packages with OIDC trusted publishing

## 🏷️ Issue Labels

- `bug` - Something isn't working
- `enhancement` - New feature or improvement
- `documentation` - Documentation improvements
- `good first issue` - Good for newcomers
- `help wanted` - Extra attention needed  
- `excel-com` - Excel COM automation issues
- `power-query` - Power Query specific
- `coding-agent` - Coding agent related

---

Thank you for contributing to Sbroenne.ExcelMcp! Together we're making Excel automation more accessible to coding agents and developers worldwide. 🚀
