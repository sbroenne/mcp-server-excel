# ExcelMcp Test Guide

This document explains how to run different types of tests in the ExcelMcp project.

## Test Categories

### Unit Tests (Fast, Default)

- **What**: Argument validation, string matching, error handling logic
- **Requirements**: No Excel installation needed
- **Speed**: Very fast (< 1 second)
- **Run by default**: Yes

### Integration Tests (Medium Speed, Default)

- **What**: Excel COM operations, Power Query, VBA, file operations
- **Requirements**: Excel installation + Windows
- **Speed**: Medium (5-15 seconds)
- **Run by default**: Yes

### Round Trip Tests (Slow, On-Request Only)

- **What**: Complex end-to-end workflows combining multiple ExcelMcp features
- **Requirements**: Excel installation + Windows
- **Speed**: Very slow (30+ seconds)
- **Run by default**: No - only when explicitly requested

## Running Tests

### Default (Unit Tests Only in CI)

```bash
# Runs only unit tests (no Excel required) - safe for CI environments
dotnet test --filter "Category=Unit"

# Local development: runs unit tests and integration tests (requires Excel)
dotnet test
```

### By Category

```bash
# Run only unit tests (no Excel required)
dotnet test --filter "Category=Unit"

# Run only integration tests (requires Excel)
dotnet test --filter "Category=Integration"

# Run only round trip tests (requires Excel, slow)
dotnet test --filter "Category=RoundTrip"
```

### By Speed

```bash
# Run only fast tests
dotnet test --filter "Speed=Fast"

# Run fast and medium speed tests (exclude slow)
dotnet test --filter "Speed=Fast|Speed=Medium"
```

### By Feature

```bash
# Run only PowerQuery tests
dotnet test --filter "Feature=PowerQuery"

# Run only VBA script tests
dotnet test --filter "Feature=VBA"

# Run only worksheet tests
dotnet test --filter "Feature=Worksheets"

# Run only file operation tests
dotnet test --filter "Feature=Files"
```

### Specific Test Classes

```bash
# Run only PowerQuery integration tests
dotnet test --filter "FullyQualifiedName~PowerQueryCommandsTests"

# Run only unit tests
dotnet test --filter "FullyQualifiedName~UnitTests"
```

## CI/CD Considerations

### GitHub Actions / Azure DevOps (No Excel Available)

```yaml
# CI environments typically don't have Excel installed
- name: Run Unit Tests
  run: dotnet test --filter "Category=Unit"
```

### Self-Hosted Runners with Excel (Optional)

```yaml
# Only if you have Windows runners with Excel installed
- name: Run Integration Tests
  run: dotnet test --filter "Category=Integration"
  # This requires Windows runners with Excel installation

- name: Run Round Trip Tests
  run: dotnet test --filter "Category=RoundTrip"
  # This requires Windows runners with Excel installation
```

### Local Development

```bash
# Quick feedback loop during development (unit tests only)
dotnet test --filter "Category=Unit"

# Feature testing with Excel (integration tests)
dotnet test --filter "Category=Integration"

# Full validation including slow round trip tests
dotnet test --filter "Category=RoundTrip"

# All non-slow tests (unit + integration)
dotnet test --filter "Speed!=Slow"
```

## Test Structure

```text
tests/
├── ExcelMcp.Tests/
│   ├── UnitTests.cs                     # [Unit, Fast] - No Excel required
│   └── Commands/
│       ├── FileCommandsTests.cs        # [Integration, Medium, Files] - Excel file operations
│       ├── PowerQueryCommandsTests.cs  # [Integration, Medium, PowerQuery] - M code automation
│       ├── ScriptCommandsTests.cs      # [Integration, Medium, VBA] - VBA script operations
│       ├── SheetCommandsTests.cs       # [Integration, Medium, Worksheets] - Sheet operations
│       └── IntegrationRoundTripTests.cs # [RoundTrip, Slow, EndToEnd] - Complex workflows
```

## Test Organization in Test Explorer

Tests are organized using multiple traits for better filtering:

- **Category**: `Unit`, `Integration`, `RoundTrip`
- **Speed**: `Fast`, `Medium`, `Slow`
- **Feature**: `PowerQuery`, `VBA`, `Worksheets`, `Files`, `EndToEnd`

## Environment Requirements

### Unit Tests (`Category=Unit`)

- **Requirements**: None
- **Platforms**: Windows, Linux, macOS
- **CI Compatible**: ✅ Yes
- **Purpose**: Validate argument parsing, logic, validation

### Integration Tests (`Category=Integration`)

- **Requirements**: Windows + Excel installation
- **Platforms**: Windows only
- **CI Compatible**: ❌ No (unless using Windows runners with Excel)
- **Purpose**: Validate Excel COM operations, feature functionality

### Round Trip Tests (`Category=RoundTrip`)

- **Requirements**: Windows + Excel installation + VBA trust settings
- **Platforms**: Windows only
- **CI Compatible**: ❌ No (unless using specialized Windows runners)
- **Purpose**: End-to-end workflow validation

## Troubleshooting

### "Round trip tests skipped" Message

This is expected behavior. Round trip tests only run when explicitly requested:

- Use `dotnet test --filter "Category=RoundTrip"` to run them specifically
- Round trip tests are slow and not needed for regular development

### Excel COM Errors in Integration/Round Trip Tests

- **CI Environment**: Integration and Round Trip tests will fail without Excel
  - Use `--filter "Category=Unit"` in CI pipelines
  - Only run Excel-dependent tests on local machines or Windows runners with Excel
- Ensure Excel is installed (Windows only)
- Close all Excel instances before running tests
- Run `ExcelMcp setup-vba-trust` for VBA tests
- Excel COM is not available on Linux/macOS

### Slow Test Performance

- **Unit tests**: Very fast (< 1 second)
- **Integration tests**: Medium speed (5-15 seconds)
- **Round trip tests**: Very slow (30+ seconds)
- **CI Strategy**: Run only unit tests in CI (no Excel required)
- **Local Development**:
  - Use unit tests for rapid development cycles
  - Use integration tests for feature validation
  - Use round trip tests for comprehensive end-to-end validation

## Adding New Tests

### Fast Unit Test

```csharp
[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
public class MyUnitTests
{
    [Fact]
    public void MyMethod_WithValidInput_ReturnsExpected()
    {
        // Test logic without Excel COM
    }
}
```

### Integration Test

```csharp
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "YourFeature")]  // e.g., "PowerQuery", "VBA", "Worksheets", "Files"
public class MyCommandsTests : IDisposable
{
    [Fact]
    public void MyCommand_WithExcel_WorksCorrectly()
    {
        // Test logic using Excel COM automation
    }
}
```

### Round Trip Test

```csharp
[Trait("Category", "RoundTrip")]
[Trait("Speed", "Slow")]
[Trait("Feature", "EndToEnd")]
public class MyRoundTripTests : IDisposable
{
    [Fact]
    public void ComplexWorkflow_EndToEnd_WorksCorrectly()
    {
        // Complex workflow testing multiple features together
    }
}
```

## Benefits of This Test Organization

✅ **Fast feedback during development** (unit tests)  
✅ **Feature validation with Excel** (integration tests)  
✅ **Comprehensive end-to-end validation when requested** (round trip tests)  
✅ **Flexible filtering by category, speed, or feature**  
✅ **Better organization in Test Explorer**  
✅ **CI/CD flexibility** (different test suites for different scenarios)  
✅ **Clear documentation for contributors**
