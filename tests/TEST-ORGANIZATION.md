# Test Organization

## Overview

Tests are organized by layer to match the separation of concerns in the architecture:

```
tests/
├── ExcelMcp.Core.Tests/      ← Most tests here (data layer)
├── ExcelMcp.CLI.Tests/        ← Minimal tests (presentation layer)
└── ExcelMcp.McpServer.Tests/  ← MCP protocol tests
```

## Test Distribution

### ExcelMcp.Core.Tests (Primary Test Suite)
**Purpose**: Test the data layer - Core business logic without UI concerns

**What to test**:
- ✅ Result objects returned correctly
- ✅ Data validation logic
- ✅ Excel COM operations
- ✅ Error handling and edge cases
- ✅ File operations
- ✅ Data transformations

**Characteristics**:
- Tests call Core commands directly
- No UI concerns (no console output testing)
- Verifies Result object properties
- Most comprehensive test coverage
- **This is where 80-90% of tests should be**

**Example**:
```csharp
[Fact]
public void CreateEmpty_WithValidPath_ReturnsSuccessResult()
{
    // Arrange
    var commands = new FileCommands();
    
    // Act
    var result = commands.CreateEmpty("test.xlsx");
    
    // Assert
    Assert.True(result.Success);
    Assert.Equal("create-empty", result.Action);
    Assert.Null(result.ErrorMessage);
}
```

### ExcelMcp.CLI.Tests (Minimal Test Suite)
**Purpose**: Test CLI-specific behavior - argument parsing, exit codes, user interaction

**What to test**:
- ✅ Command-line argument parsing
- ✅ Exit codes (0 for success, 1 for error)
- ✅ User prompt handling
- ✅ Console output formatting (optional)

**Characteristics**:
- Tests call CLI commands with `string[] args`
- Verifies int return codes
- Minimal coverage - only CLI-specific behavior
- **This is where 10-20% of tests should be**

**Example**:
```csharp
[Fact]
public void CreateEmpty_WithValidPath_ReturnsZeroAndCreatesFile()
{
    // Arrange
    string[] args = { "create-empty", "test.xlsx" };
    var commands = new FileCommands();
    
    // Act
    int exitCode = commands.CreateEmpty(args);
    
    // Assert
    Assert.Equal(0, exitCode);
}
```

### ExcelMcp.McpServer.Tests
**Purpose**: Test MCP protocol compliance and JSON responses

**What to test**:
- ✅ JSON serialization correctness
- ✅ MCP tool interfaces
- ✅ Error responses in JSON format
- ✅ Protocol compliance

## Test Categories and Traits

All tests should use traits for filtering:

```csharp
[Trait("Category", "Integration")]  // Unit, Integration, RoundTrip
[Trait("Speed", "Fast")]            // Fast, Medium, Slow
[Trait("Feature", "Files")]         // Files, PowerQuery, VBA, etc.
[Trait("Layer", "Core")]            // Core, CLI, MCP
```

## Running Tests

```bash
# Run all Core tests (primary suite)
dotnet test --filter "Layer=Core"

# Run all CLI tests (minimal suite)
dotnet test --filter "Layer=CLI"

# Run fast tests only
dotnet test --filter "Speed=Fast"

# Run specific feature tests
dotnet test --filter "Feature=Files&Layer=Core"

# Run all tests except slow ones
dotnet test --filter "Speed!=Slow"
```

## Test Structure Guidelines

### Core Tests Should:
1. Test Result objects, not console output
2. Verify all properties of Result objects
3. Test edge cases and error conditions
4. Be comprehensive - this is the primary test suite
5. Use descriptive test names that explain what's being verified

### CLI Tests Should:
1. Focus on argument parsing
2. Verify exit codes
3. Be minimal - just verify CLI wrapper works
4. Not duplicate Core logic tests

### MCP Tests Should:
1. Verify JSON structure
2. Test protocol compliance
3. Verify error responses

## Migration Path

When refactoring a command type:

1. **Create Core.Tests first**:
   ```
   tests/ExcelMcp.Core.Tests/Commands/MyCommandTests.cs
   ```
   - Comprehensive tests for all functionality
   - Test Result objects

2. **Create minimal CLI.Tests**:
   ```
   tests/ExcelMcp.CLI.Tests/Commands/MyCommandTests.cs
   ```
   - Just verify argument parsing and exit codes
   - 3-5 tests typically sufficient

3. **Update MCP.Tests if needed**:
   ```
   tests/ExcelMcp.McpServer.Tests/Tools/MyToolTests.cs
   ```
   - Verify JSON responses

## Example: FileCommands Test Coverage

### Core.Tests (Comprehensive - 13 tests)
- ✅ CreateEmpty_WithValidPath_ReturnsSuccessResult
- ✅ CreateEmpty_WithNestedDirectory_CreatesDirectoryAndReturnsSuccess
- ✅ CreateEmpty_WithEmptyPath_ReturnsErrorResult
- ✅ CreateEmpty_WithRelativePath_ConvertsToAbsoluteAndReturnsSuccess
- ✅ CreateEmpty_WithValidExtensions_ReturnsSuccessResult (Theory: 2 cases)
- ✅ CreateEmpty_WithInvalidExtensions_ReturnsErrorResult (Theory: 3 cases)
- ✅ CreateEmpty_WithInvalidPath_ReturnsErrorResult
- ✅ CreateEmpty_MultipleTimes_ReturnsSuccessForEachFile
- ✅ CreateEmpty_FileAlreadyExists_WithoutOverwrite_ReturnsError
- ✅ CreateEmpty_FileAlreadyExists_WithOverwrite_ReturnsSuccess
- ✅ Validate_ExistingValidFile_ReturnsValidResult
- ✅ Validate_NonExistentFile_ReturnsInvalidResult
- ✅ Validate_FileWithInvalidExtension_ReturnsInvalidResult

### CLI.Tests (Minimal - 4 tests)
- ✅ CreateEmpty_WithValidPath_ReturnsZeroAndCreatesFile
- ✅ CreateEmpty_WithMissingArguments_ReturnsOneAndDoesNotCreateFile
- ✅ CreateEmpty_WithInvalidExtension_ReturnsOneAndDoesNotCreateFile
- ✅ CreateEmpty_WithValidExtensions_ReturnsZero (Theory: 2 cases)

### Ratio: ~77% Core, ~23% CLI
This matches the principle that most tests should focus on Core data logic.

## Benefits of This Organization

1. **Clear Separation**: Tests match the layered architecture
2. **Fast Feedback**: Core tests run without CLI overhead
3. **Better Coverage**: Comprehensive Core tests catch more bugs
4. **Easier Maintenance**: Changes to CLI formatting don't break Core tests
5. **Reusability**: Core tests work even if we add new presentation layers (web, desktop, etc.)

## Anti-Patterns to Avoid

❌ **Don't**: Put all tests in CLI.Tests
- Makes tests fragile to UI changes
- Mixes concerns
- Harder to reuse Core in other contexts

❌ **Don't**: Test console output in Core.Tests
- Core shouldn't have console output
- Tests should verify Result objects, not strings

❌ **Don't**: Duplicate Core logic tests in CLI.Tests
- CLI tests should be minimal
- Core tests already cover the logic

✅ **Do**: Put most tests in Core.Tests
✅ **Do**: Test Result objects in Core.Tests
✅ **Do**: Keep CLI.Tests minimal and focused on presentation
✅ **Do**: Use traits to organize and filter tests
