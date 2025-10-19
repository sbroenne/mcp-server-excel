# Architecture Refactoring: Separation of Concerns

## Overview

This document describes the refactoring of ExcelMcp to separate the **data layer (Core)** from the **presentation layer (CLI/MCP Server)**.

## Problem Statement

Previously, the Core library mixed data operations with console formatting using Spectre.Console:
- Core commands returned `int` (0=success, 1=error)
- Core commands directly wrote to console with `AnsiConsole.MarkupLine()`
- MCP Server and CLI both depended on Core's output format
- Core could not be used in non-console scenarios

## New Architecture

### Core Layer (Data-Only)
**Purpose**: Pure data operations, no formatting, no console I/O

**Characteristics**:
- Returns strongly-typed Result objects (OperationResult, FileValidationResult, etc.)
- No Spectre.Console dependency
- No console output
- No user prompts
- Focuses on Excel COM interop and data operations

**Example - FileCommands.CreateEmpty**:
```csharp
// Returns structured data, not console output
public OperationResult CreateEmpty(string filePath, bool overwriteIfExists = false)
{
    // ... Excel operations ...
    
    return new OperationResult
    {
        Success = true,
        FilePath = filePath,
        Action = "create-empty"
    };
}
```

### CLI Layer (Console Formatting)
**Purpose**: Wrap Core commands and format results for console users

**Characteristics**:
- Uses Spectre.Console for rich console output
- Handles user prompts and confirmations
- Calls Core commands and formats the Result objects
- Maintains `string[] args` interface for backward compatibility

**Example - CLI FileCommands**:
```csharp
public int CreateEmpty(string[] args)
{
    // Parse arguments and handle user interaction
    bool overwrite = File.Exists(filePath) && 
                     AnsiConsole.Confirm("Overwrite?");
    
    // Call Core (no formatting)
    var result = _coreCommands.CreateEmpty(filePath, overwrite);
    
    // Format output for console
    if (result.Success)
    {
        AnsiConsole.MarkupLine($"[green]✓[/] Created: {filePath}");
        return 0;
    }
    else
    {
        AnsiConsole.MarkupLine($"[red]Error:[/] {result.ErrorMessage}");
        return 1;
    }
}
```

### MCP Server Layer (JSON Output)
**Purpose**: Expose Core commands as JSON API for AI clients

**Characteristics**:
- Calls Core commands directly
- Serializes Result objects to JSON
- Optimized for MCP protocol clients (Claude, ChatGPT, GitHub Copilot)
- No console formatting

**Example - MCP Server ExcelTools**:
```csharp
private static string CreateEmptyFile(FileCommands fileCommands, 
                                      string filePath, 
                                      bool macroEnabled)
{
    var result = fileCommands.CreateEmpty(filePath, overwriteIfExists: false);
    
    // Return clean JSON for MCP clients
    return JsonSerializer.Serialize(new
    {
        success = result.Success,
        filePath = result.FilePath,
        macroEnabled,
        message = result.Success ? "Excel file created successfully" : null,
        error = result.ErrorMessage
    });
}
```

## Benefits

### 1. **Separation of Concerns**
- Core: Pure data logic
- CLI: Console user experience
- MCP Server: JSON API for AI clients

### 2. **Reusability**
Core can now be used in:
- Console applications (CLI)
- AI assistants (MCP Server)
- Web APIs
- Desktop applications
- Unit tests (easier to test data operations)

### 3. **Maintainability**
- Changes to console formatting don't affect Core
- Changes to JSON format don't affect Core
- Core logic can be tested independently

### 4. **Testability**
Tests can verify Result objects instead of parsing console output:
```csharp
// Before: Hard to test
int result = command.CreateEmpty(args);
Assert.Equal(0, result); // Only knows success/failure

// After: Easy to test
var result = command.CreateEmpty(filePath);
Assert.True(result.Success);
Assert.Equal("create-empty", result.Action);
Assert.Equal(expectedPath, result.FilePath);
Assert.Null(result.ErrorMessage);
```

## Migration Status

### ✅ Completed
- **FileCommands**: Fully refactored
  - Core returns `OperationResult` and `FileValidationResult`
  - CLI wraps Core and formats with Spectre.Console
  - MCP Server returns clean JSON
  - All tests updated

### 🔄 Remaining Work
The same pattern needs to be applied to:
- PowerQueryCommands → `PowerQueryListResult`, `PowerQueryViewResult`
- SheetCommands → `WorksheetListResult`, `WorksheetDataResult`
- ParameterCommands → `ParameterListResult`, `ParameterValueResult`
- CellCommands → `CellValueResult`
- ScriptCommands → `ScriptListResult`
- SetupCommands → `OperationResult`

## Result Types

All Result types are defined in `src/ExcelMcp.Core/Models/ResultTypes.cs`:

- `ResultBase` - Base class with Success, ErrorMessage, FilePath
- `OperationResult` - For create/delete/update operations
- `FileValidationResult` - For file validation
- `WorksheetListResult` - For listing worksheets
- `WorksheetDataResult` - For reading worksheet data
- `PowerQueryListResult` - For listing Power Queries
- `PowerQueryViewResult` - For viewing Power Query code
- `ParameterListResult` - For listing named ranges
- `ParameterValueResult` - For parameter values
- `CellValueResult` - For cell operations
- `ScriptListResult` - For VBA scripts

## Implementation Guidelines

When refactoring a command:

1. **Update Core Interface**:
   ```csharp
   // Change from:
   int MyCommand(string[] args);
   
   // To:
   MyResultType MyCommand(string param1, string param2);
   ```

2. **Update Core Implementation**:
   - Remove all `AnsiConsole` calls
   - Return Result objects instead of int
   - Remove argument parsing (CLI's responsibility)

3. **Update CLI Wrapper**:
   - Keep `string[] args` interface
   - Parse arguments
   - Handle user prompts
   - Call Core command
   - Format Result with Spectre.Console

4. **Update MCP Server**:
   - Call Core command
   - Serialize Result to JSON

5. **Update Tests**:
   - Test Result objects instead of int return codes
   - Verify Result properties

## Example: Complete Refactoring

See `FileCommands` for a complete example:
- Core: `src/ExcelMcp.Core/Commands/FileCommands.cs`
- CLI: `src/ExcelMcp.CLI/Commands/FileCommands.cs`
- MCP: `src/ExcelMcp.McpServer/Tools/ExcelTools.cs` (CreateEmptyFile method)
- Tests: `tests/ExcelMcp.CLI.Tests/Commands/FileCommandsTests.cs`

## Backward Compatibility

CLI interface remains unchanged:
```bash
# Still works the same way
excelcli create-empty myfile.xlsx
```

Users see no difference in CLI behavior, but the architecture is cleaner.

## Future Enhancements

With this architecture, we can easily:
1. Add web API endpoints
2. Create WPF/WinForms UI
3. Build VS Code extension
4. Add gRPC server
5. Create REST API

All by reusing the Core data layer and adding new presentation layers.
