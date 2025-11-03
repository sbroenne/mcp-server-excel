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
   - .NET 8 SDK
   - Microsoft Excel installed

2. **Setup**:
   ```powershell
   git clone https://github.com/sbroenne/mcp-server-excel.git
   cd ExcelMcp
   dotnet restore
   dotnet build
   ```

## 🚨 **CRITICAL: Pull Request Workflow Required**

**All changes must be made through Pull Requests (PRs).** Direct commits to `main` are prohibited.

### Quick PR Process

1. **Create feature branch**: `git checkout -b feature/your-feature`
2. **Make changes**: Code, tests, documentation
3. **Push branch**: `git push origin feature/your-feature`
4. **Create PR**: Use GitHub's PR template
5. **Address review**: Make requested changes
6. **Merge**: After approval and CI checks pass

📋 **Detailed workflow**: See [DEVELOPMENT.md](DEVELOPMENT.md) for complete instructions.

3. **Test Your Setup**:
   ```powershell
   dotnet run -- pq-list "path/to/test.xlsx"
   ```

## 📋 Development Guidelines

### Code Style

- **C# 12** features encouraged (file-scoped namespaces, records, pattern matching)
- **Nullable reference types** enabled - handle nulls properly
- **No warnings** - project must build with zero warnings
- **XML documentation** for public APIs
- **Consistent naming** - follow established patterns

### Architecture Patterns

#### Command Pattern
All commands follow this structure:

```csharp
// Interface
public interface IMyCommands
{
    int MyOperation(string[] args);
}

// Implementation  
public class MyCommands : IMyCommands
{
    public int MyOperation(string[] args)
    {
        // Validation
        if (!ValidateArgs(args, expectedCount, "usage string"))
            return 1;
            
        // Excel automation using helper
        return ExcelHelper.WithExcel(filePath, save, (excel, workbook) =>
        {
            // Your implementation
            return 0; // Success
        });
    }
}
```

#### Critical Rules

1. **Always use `ExcelHelper.WithExcel()`** - Never manage Excel lifecycle manually
2. **Excel uses 1-based indexing** - `collection.Item(1)` is the first element
3. **Use `QueryTables.Add()` not `ListObjects.Add()`** - For loading Power Query data
4. **Escape user input** - Always use `.EscapeMarkup()` with Spectre.Console
5. **Return 0 for success, 1+ for errors** - Consistent exit codes

### Excel COM Best Practices

- **Late binding with dynamic types** - Use `Type.GetTypeFromProgID("Excel.Application")`
- **Proper error handling** - Catch `COMException` and provide helpful messages
- **Resource cleanup** - Let `ExcelHelper` handle COM object lifecycle
- **Input validation** - Check file existence and argument counts early

### Testing

Before submitting:

1. **Manual testing** with various Excel files
2. **Verify Excel process cleanup** - No `excel.exe` should remain after 5 seconds
3. **Test error conditions** - Missing files, invalid arguments, etc.
4. **VBA script testing** - For script-related commands, test with real VBA macros
5. **Cross-version compatibility** - Test with different Excel versions if possible

## 🔧 Adding New Commands

### 1. Create Interface

```csharp
// Commands/INewCommands.cs
namespace ExcelMcp.Commands;

public interface INewCommands
{
    int NewOperation(string[] args);
}
```

### 2. Implement Command Class

```csharp
// Commands/NewCommands.cs
using Spectre.Console;

namespace ExcelMcp.Commands;

public class NewCommands : INewCommands
{
    public int NewOperation(string[] args)
    {
        // Implementation following established patterns
    }
}
```

### 3. Register in Program.cs

Add to the switch expression in `Main()`:

```csharp
return args[0] switch
{
    "new-operation" => newCommands.NewOperation(args),
    // ... existing commands
    _ => ShowHelp()
};
```

### 4. Update Help Text

Add your command to the help output in `ShowHelp()`.

## 📝 Pull Request Process

### Before Submitting

- [ ] Code builds with zero warnings
- [ ] All existing commands still work
- [ ] Excel processes clean up properly
- [ ] Added appropriate error handling
- [ ] Updated help text if needed
- [ ] Tested with various Excel files

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
