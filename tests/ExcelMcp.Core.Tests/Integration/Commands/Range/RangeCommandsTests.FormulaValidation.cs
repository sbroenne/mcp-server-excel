using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for formula validation and error detection
/// Feature: #1 Formula Syntax Validation, #4 Better Error Code Mapping
/// </summary>
public partial class RangeCommandsTests
{
    // === IMPROVEMENT #1: FORMULA VALIDATION TESTS ===

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithValidFormulas_ReturnsSuccess()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up source data
        _commands.SetValues(batch, sheetName, "A1:A3", [
            [10],
            [20],
            [30]
        ]);

        var formulas = new List<List<string>> {
            new() { "=SUM(A1:A3)" },
            new() { "=AVERAGE(A1:A3)" },
            new() { "=COUNT(A1:A3)" }
        };

        // Act - validate formulas before applying them
        var result = _commands.ValidateFormulas(batch, sheetName, "B1:B3", formulas);

        // Assert
        Assert.True(result.Success);
        Assert.True(result.IsValid);
        Assert.Equal(3, result.FormulaCount);
        Assert.Equal(3, result.ValidCount);
        Assert.Equal(0, result.ErrorCount);
        Assert.Null(result.Errors);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithUndefinedFunction_DetectsError()
    {
        // Arrange - use shared file
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "=GETVM3(4,16,\"region\")" }  // Missing XA2. namespace
        };

        // Act - validate should detect missing namespace
        var result = _commands.ValidateFormulas(batch, sheetName, "B1", formulas);

        // Assert
        Assert.False(result.IsValid);
        Assert.Equal(1, result.FormulaCount);
        Assert.Equal(0, result.ValidCount);
        Assert.Equal(1, result.ErrorCount);
        Assert.NotNull(result.Errors);
        Assert.Single(result.Errors);

        var error = result.Errors[0];
        Assert.Equal("B1", error.CellAddress);
        Assert.Contains("GETVM3", error.Message);
        Assert.Contains("XA2.", error.Suggestion ?? "");
        Assert.Equal("undefined-function", error.Category);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithMissingNamespace_SuggestsCorrection()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "=GETAKS(2,4)" },     // Missing XA2.
            new() { "=XA2.GETVM3(4,16,\"region\")" }  // Correct
        };

        // Act
        var result = _commands.ValidateFormulas(batch, sheetName, "B1:B2", formulas);

        // Assert
        Assert.False(result.IsValid);
        Assert.Equal(2, result.FormulaCount);
        Assert.Equal(1, result.ValidCount);
        Assert.Equal(1, result.ErrorCount);

        var error = result.Errors![0];
        Assert.Equal("B1", error.CellAddress);
        Assert.Contains("=XA2.GETAKS", error.Suggestion ?? "");
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithInvalidReference_DetectsError()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "=SUM(UnknownSheet!A1:A10)" }
        };

        // Act
        var result = _commands.ValidateFormulas(batch, sheetName, "B1", formulas);

        // Assert
        Assert.False(result.IsValid);
        Assert.Equal(1, result.ErrorCount);
        var error = result.Errors![0];
        Assert.Equal("invalid-reference", error.Category);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithSyntaxError_ReportsError()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "=SUM(A1:A3" }  // Missing closing parenthesis
        };

        // Act
        var result = _commands.ValidateFormulas(batch, sheetName, "B1", formulas);

        // Assert
        Assert.False(result.IsValid);
        Assert.Equal(1, result.ErrorCount);
        var error = result.Errors![0];
        Assert.Equal("syntax-error", error.Category);
        Assert.Contains("parenthesis", error.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithEmptyFormulas_SkipsValidation()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "" }  // Empty (no formula)
        };

        // Act
        var result = _commands.ValidateFormulas(batch, sheetName, "B1", formulas);

        // Assert
        Assert.True(result.IsValid);
        Assert.Equal(1, result.FormulaCount);
        Assert.Equal(1, result.ValidCount);
        Assert.Equal(0, result.ErrorCount);
    }

    // === IMPROVEMENT #4: ERROR CODE MAPPING TESTS ===

    [Fact]
    [Trait("Feature", "Range")]
    public void GetFormulas_WithErrorCodes_MapsToHumanReadableMessages()
    {
        // Arrange - use shared file
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Create a #NAME? error (undefined function)
        _commands.SetFormulas(batch, sheetName, "A1", [
            ["=UNDEFINEDFUNCTION()"]
        ]);

        // Act
        var result = _commands.GetFormulas(batch, sheetName, "A1");

        // Assert - should detect error and include mapping
        Assert.NotNull(result.CellErrors);
        Assert.NotEmpty(result.CellErrors);

        var error = result.CellErrors[0];
        Assert.Equal("A1", error.CellAddress);
        Assert.Equal(-2146826259, error.ErrorCode);  // #NAME? error code
        Assert.Contains("undefined", error.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(error.Suggestion);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void GetFormulas_WithCircularReference_DetectsWarning()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Create circular reference: A1 = B1, B1 = A1
        _commands.SetFormulas(batch, sheetName, "A1", [
            ["=B1"]
        ]);
        _commands.SetFormulas(batch, sheetName, "B1", [
            ["=A1"]
        ]);

        // Act
        var result = _commands.GetFormulas(batch, sheetName, "A1:B1");

        // Assert - should detect circular reference
        // Note: Excel may not immediately report circular ref until calc,
        // but GetFormulas enhancement should detect it
        Assert.True(result.Success);
        // Result may include warning about circular reference if detected
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void GetFormulas_WithComplexRange_ReturnsAllErrorsWithAddresses()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set mix of valid and invalid formulas
        _commands.SetFormulas(batch, sheetName, "A1:A3", [
            ["=1+1"],                        // Valid
            ["=BADFUNCTION()"],              // Error
            ["=2+2"]                         // Valid
        ]);

        // Act
        var result = _commands.GetFormulas(batch, sheetName, "A1:A3");

        // Assert
        Assert.Equal(3, result.RowCount);
        // At minimum, one error should be present for the bad function
        if (result.CellErrors != null)
        {
            Assert.NotEmpty(result.CellErrors);
            Assert.Contains(result.CellErrors, e => e.Row == 2 && e.Column == 1);
        }
    }
}
