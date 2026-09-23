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
3. **Run the pre-commit hook**: follow the
   [pre-commit setup guide](PRE-COMMIT-SETUP.md), then let it run on every
   commit. It checks COM cleanup, MCP/CLI parity, the Release build, packaging,
   smoke tests, and other required gates. Never bypass it with `--no-verify`.
4. **Push branch**: `git push origin feature/your-feature`
5. **Create PR**: Use GitHub's PR template
6. **Address review**: Investigate human and automated comments, fix verified defects, and explain why an incorrect or inapplicable suggestion was not applied. Do not make unrelated style changes simply because a bot suggested them.
7. **Merge**: After approval and CI checks pass — **GitHub will squash commits automatically**
   - Verify the final commit message accurately describes the changes
   - After merge, your feature branch can be safely deleted

📋 **Detailed workflow**: See [DEVELOPMENT.md](DEVELOPMENT.md) for complete instructions.

## 📋 Development Guidelines

### Portable npm lockfiles

Every npm project with a tracked lockfile needs its own `.npmrc` containing:

```ini
omit-lockfile-registry-resolved=true
```

This applies to the repository root, `vscode-extension`, and
`videos/excel-mcp-intro`, as well as any new nested npm project. npm does not
inherit the root project's configuration in nested projects. Preserve any
other existing project settings; do not change registry, proxy, credentials,
or user/global npm configuration to clean up a lockfile.

Run this command **inside each affected project**:

```powershell
npm install --package-lock-only --ignore-scripts
```

Use npm to regenerate lockfiles, not search-and-replace. For portability-only
changes, keep dependency versions and integrity hashes unchanged; do not run
`npm update` or `npm audit fix`. Downloads then use the developer's or CI
environment's configured registry rather than a URL saved by another machine.
Direct URL dependencies are not portable under this policy and must use
registry versions or local dependencies instead.

Run the focused checks from the repository root (no Excel required):

```powershell
pwsh -NoProfile -File scripts\Test-NpmLockfiles.ps1
pwsh -NoProfile -File scripts\check-npm-lockfiles.ps1
```

The required CI Gate runs both checks with two-minute limits. The pre-commit
hook checks the staged contents with `-Staged`. The guard discovers tracked
`package-lock.json` and `npm-shrinkwrap.json` files at any depth, excludes
`node_modules`, and reports offending filenames without exposing URLs or
credentials. New lockfiles must be staged before the guard can discover them.

### Respect local package sources

npm installs and .NET restores use the package manager's normal configuration
hierarchy. The repository does not clear NuGet package sources or force a
registry/source/configuration file for restores. Keep local feeds, credentials,
proxy settings, caches, and environment settings under the developer's or CI
environment's control; do not overwrite them to work around a restore failure.
Report an unavailable configured feed instead.

NuGet publishing commands intentionally name the public publishing destination.
That `dotnet nuget push --source` setting is not a restore-source override.

### Code Style

- **C# version** follows `Directory.Build.props` and the SDK selected by `global.json`
- **Nullable reference types** enabled - handle nulls properly
- **No warnings** - project must build with zero warnings
- **XML documentation** for public APIs (these docs are extracted into MCP tool descriptions and shown to LLMs — keep them accurate)
- **Consistent naming** - follow established patterns
- **Type organization** - one public type per file, with a matching file name;
  split large command classes into domain-specific partial files
- **Typed boundaries** - use result models for cross-layer data rather than
  anonymous or loosely typed payloads

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
public OperationResult Rename(IExcelBatch batch, string oldName, string newName)
{
    ArgumentException.ThrowIfNullOrWhiteSpace(oldName);
    ArgumentException.ThrowIfNullOrWhiteSpace(newName);

    return batch.Execute((ctx, ct) =>
    {
        Excel.Worksheet? sheet = null;
        try
        {
            ct.ThrowIfCancellationRequested();
            sheet = ComUtilities.FindSheet(ctx.Book, oldName)
                ?? throw new InvalidOperationException(
                    $"Worksheet '{oldName}' was not found.");
            sheet.Name = newName;
            return new OperationResult { Success = true };
        }
        finally
        {
            ComUtilities.Release(ref sheet);
        }
    });
}
```

Here `Excel` aliases `Microsoft.Office.Interop.Excel`. Validate ordinary .NET
arguments before entering the batch. The batch propagates callback failures to
the caller; Service and MCP boundaries serialize failures with their diagnostic
context. Do not replace that context with a second generic error result.

#### Critical Rules

1. **Always use the batch API** - Never manage Excel lifecycle manually
2. **Excel uses 1-based indexing** - `collection.Item(1)` is the first element
3. **Never suppress exceptions** with a catch block that returns `Success = false` — let `batch.Execute()` handle it
4. **`Success = true` must never coexist with a non-empty `ErrorMessage`**
5. **COM objects** are released only in `finally` blocks, never swallowed in empty `catch` blocks

### Excel COM Best Practices

- **Typed Excel PIAs first** - use late binding only for documented PIA/runtime dependency gaps
- **Proper error handling** - Catch `COMException` where specific handling is needed; otherwise let exceptions propagate
- **Resource cleanup** - the batch owns its application and workbook; release every COM reference acquired by a command in reverse order in `finally`, including intermediate collections
- **Input validation** - Check file existence and argument validity early
- **Performance** - reuse sessions and bulk range operations instead of per-cell COM calls

See the [COM pitfalls](../.github/instructions/excel-com-interop.instructions.md)
for application-state, refresh, numeric conversion, and shutdown constraints.

### Testing

COM behavior requires real Excel integration tests. Pure parsing, mapping,
serialization, and generation can use focused tests without Excel. For behavior
changes, write a failing regression test before implementation. See the
[test guide](../tests/README.md) for fixtures, assertions, and persistence.

```powershell
# Select the affected project and feature; use a hard execution timeout
dotnet test tests\ExcelMcp.Core.Tests\ExcelMcp.Core.Tests.csproj --filter "Feature=PowerQuery&RunType!=OnDemand"

# Session/batch changes also require relevant ComInterop OnDemand tests
dotnet test tests\ExcelMcp.ComInterop.Tests\ExcelMcp.ComInterop.Tests.csproj --filter "RunType=OnDemand"
```

Before submitting a PR:

1. Tests pass for the feature(s) you changed
2. Test-owned Excel processes clean up; do not terminate unrelated user sessions
3. Error conditions tested (missing files, invalid arguments, etc.)
4. Build has zero warnings
5. Pre-commit hook passes every gate applicable to the staged paths. Excel E2E is required only when Core, CLI, or MCP runtime paths change, including `ComInterop`, `Service`, and their source generators.

## 🔧 Adding a New Operation

New operations are added to the **Core** interface/implementation; CLI commands and MCP tool schemas are then generated automatically — you don't hand-write CLI arg parsing or MCP tool registration.

1. **Add the method to the relevant Core interface** (e.g. `Commands/Sheet/ISheetCommands.cs`), with XML doc comments (these become the MCP tool/parameter descriptions).
2. **Implement it** in the corresponding partial class (e.g. `SheetCommands.Lifecycle.cs`), following the batch-API pattern above.
3. **Build the solution** - the source generators (`ExcelMcp.Generators`, `ExcelMcp.Generators.CLI`) produce the CLI verb and MCP tool automatically from the interface.
4. **Add integration tests** for the new operation (TDD: write them first).
5. **Update `FEATURES.md` and the appropriate `docs/features/*.md` file** with the new operation and updated operation count — `scripts/check-doc-counts.ps1` enforces that documented counts match the code.

### Tracing a bug or contract change

Start at the failing entry point and trace generated routing, Service, Core,
and Excel to identify the owning layer. Check sibling operations, fallback and
retry branches, and cached or parallel paths for the same defect before choosing
a fix.

For changed actions or parameters, compare the Core contract, generated Service
arguments, CLI options and batch JSON, MCP schema and manual exceptions, tests,
and shared guidance. Names, defaults, validation, results, and timeout behavior
must agree. A successful build does not establish that every operation is exposed;
run the applicable [repository audits](../.github/copilot-instructions.md#build-and-validation).

Reproduce the bug in a focused test, observe the failure, fix the owning layer,
then rerun that test and the smallest related group. Coverage should follow the
risk, not a fixed number of tests or documentation edits.

### Documentation changes

Keep entry READMEs focused on their audience: repository acquisition and quick
start, component installation/use, or Marketplace benefits. Put detailed feature
behavior in `docs/features/` and shared agent workflows in `skills/shared/`.
There is no fixed README length or requirement to edit every README.

Before shortening or moving a page, identify where each substantive caveat,
example, installation option, and workflow will remain. Update that destination
first, then replace duplicate material with a link. Permanent guides belong in
`docs/`, decisions in `docs/ADR-*.md`, and feature requirements in `specs/`.
Temporary investigations belong in issue/PR discussions, not SUMMARY/FIX files.

Use current declared action names and verify operation tables, not just headline
counts. The count audit derives the advertised surface from generated metadata.
See the [website authoring guide](../gh-pages/README.md#publishing-canonical-documentation)
for source maps, wrappers, navigation, and machine-readable outputs.

## 📝 Pull Request Process

### Before Submitting

- [ ] Code builds with zero warnings
- [ ] Feature-scoped tests pass (`dotnet test --filter "Feature=<name>&RunType!=OnDemand"`)
- [ ] Excel processes clean up properly
- [ ] Added appropriate error handling (no suppressed exceptions)
- [ ] Updated `FEATURES.md` and `docs/features/*.md` if operation counts or behaviors changed
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
- [.NET COM Interop Guide](https://learn.microsoft.com/en-us/dotnet/framework/interop/)

## 📦 For Maintainers

- [NuGet Publishing Guide](NUGET-GUIDE.md) - Complete guide for publishing all packages with OIDC trusted publishing

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
