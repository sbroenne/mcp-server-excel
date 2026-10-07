using Xunit;
using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for formula validation and formula-error reporting
/// Feature: #1 Formula Syntax Validation, #4 Better Error Code Mapping
/// </summary>
public sealed partial class PersistentServiceRangeFormulaValidationTests
{
    // === IMPROVEMENT #1: FORMULA VALIDATION TESTS ===

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithValidFormulas_ReturnsSuccess()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up source data
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1:A3", [
            [10],
            [20],
            [30]
        ]));

        var formulas = new List<List<string>> {
            new() { "=SUM(A1:A3)" },
            new() { "=AVERAGE(A1:A3)" },
            new() { "=COUNT(A1:A3)" }
        };

        // Act - validate formulas before applying them
        var result = _commands.ValidateFormulas(batch, sheetName, "B1:B3", formulas);

        // Assert
        RequireSuccess(result);
        Assert.True(result.IsValid);
        Assert.Equal(3, result.FormulaCount);
        Assert.Equal(3, result.ValidCount);
        Assert.Equal(0, result.ErrorCount);
        Assert.Null(result.Errors);
        var values = RequireSuccess(_commands.GetValues(batch, sheetName, "B1:B3"));
        var readFormulas = RequireSuccess(_commands.GetFormulas(batch, sheetName, "B1:B3"));
        Assert.Equal(3, values.Values.Count);
        Assert.Equal(3, readFormulas.Formulas.Count);
        Assert.All(values.Values, row => Assert.Null(Assert.Single(row)));
        Assert.All(readFormulas.Formulas, row => Assert.Equal("", Assert.Single(row)));
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithUndefinedFunction_DetectsError()
    {
        // Arrange - use shared file
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "=GETVM3(4,16,\"region\")" }  // Missing XA2. namespace
        };
        RequireSuccess(_commands.SetValues(batch, sheetName, "B1", [["guard"]]));

        // Act - validate should detect missing namespace
        var result = _commands.ValidateFormulas(batch, sheetName, "B1", formulas);

        // Assert
        RequireSuccess(result);
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
        AssertCellUnchanged(batch, sheetName, "B1", "guard", null);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithMissingNamespace_SuggestsCorrection()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "=GETAKS(2,4)" },     // Missing XA2.
            new() { "=XA2.GETVM3(4,16,\"region\")" }  // Correct
        };
        RequireSuccess(_commands.SetValues(batch, sheetName, "B1:B2", [["guard-1"], ["guard-2"]]));

        // Act
        var result = _commands.ValidateFormulas(batch, sheetName, "B1:B2", formulas);

        // Assert
        RequireSuccess(result);
        Assert.False(result.IsValid);
        Assert.Equal(2, result.FormulaCount);
        Assert.Equal(1, result.ValidCount);
        Assert.Equal(1, result.ErrorCount);

        var error = result.Errors![0];
        Assert.Equal("B1", error.CellAddress);
        Assert.Contains("=XA2.GETAKS", error.Suggestion ?? "");
        AssertCellUnchanged(batch, sheetName, "B1", "guard-1", null);
        AssertCellUnchanged(batch, sheetName, "B2", "guard-2", null);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithInvalidReference_DetectsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "=SUM(UnknownSheet!A1:A10)" }
        };
        RequireSuccess(_commands.SetValues(batch, sheetName, "B1", [["guard"]]));

        // Act
        var result = _commands.ValidateFormulas(batch, sheetName, "B1", formulas);

        // Assert
        RequireSuccess(result);
        Assert.False(result.IsValid);
        Assert.Equal(1, result.ErrorCount);
        var error = result.Errors![0];
        Assert.Equal("B1", error.CellAddress);
        Assert.Equal(formulas[0][0], error.Formula);
        Assert.Equal("invalid-reference", error.Category);
        AssertCellUnchanged(batch, sheetName, "B1", "guard", null);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithSyntaxError_ReportsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "=SUM(A1:A3" }  // Missing closing parenthesis
        };
        RequireSuccess(_commands.SetValues(batch, sheetName, "B1", [["guard"]]));

        // Act
        var result = _commands.ValidateFormulas(batch, sheetName, "B1", formulas);

        // Assert
        RequireSuccess(result);
        Assert.False(result.IsValid);
        Assert.Equal(1, result.ErrorCount);
        var error = result.Errors![0];
        Assert.Equal("B1", error.CellAddress);
        Assert.Equal(formulas[0][0], error.Formula);
        Assert.Equal("syntax-error", error.Category);
        Assert.Contains("parenthesis", error.Message, StringComparison.OrdinalIgnoreCase);
        AssertCellUnchanged(batch, sheetName, "B1", "guard", null);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void ValidateFormulas_WithEmptyFormulas_SkipsValidation()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var formulas = new List<List<string>> {
            new() { "" }  // Empty (no formula)
        };
        RequireSuccess(_commands.SetValues(batch, sheetName, "B1", [["guard"]]));

        // Act
        var result = _commands.ValidateFormulas(batch, sheetName, "B1", formulas);

        // Assert
        RequireSuccess(result);
        Assert.True(result.IsValid);
        Assert.Equal(1, result.FormulaCount);
        Assert.Equal(1, result.ValidCount);
        Assert.Equal(0, result.ErrorCount);
        AssertCellUnchanged(batch, sheetName, "B1", "guard", null);
    }

    // === IMPROVEMENT #4: ERROR CODE MAPPING TESTS ===

    [Fact]
    [Trait("Feature", "Range")]
    public void GetFormulas_WithErrorCodes_MapsToHumanReadableMessages()
    {
        // Arrange - use shared file
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Current Excel versions report an unsupported function as #NAME?.
        RequireSuccess(_commands.SetFormulas(batch, sheetName, "A1", [
            ["=UNDEFINEDFUNCTION()"]
        ]));

        // Act
        var result = _commands.GetFormulas(batch, sheetName, "A1");

        // Assert - should detect error and include mapping
        RequireSuccess(result);
        Assert.NotNull(result.CellErrors);
        Assert.NotEmpty(result.CellErrors);

        var error = result.CellErrors[0];
        Assert.Equal("A1", error.CellAddress);
        Assert.Equal("#NAME?", error.ErrorName);
        Assert.Equal(-2146826259, error.ErrorCode);
        Assert.False(string.IsNullOrWhiteSpace(error.Formula));
        Assert.Contains("formula name", error.ErrorMessage, StringComparison.OrdinalIgnoreCase);
        Assert.NotNull(error.Suggestion);
        Assert.Equal("=UNDEFINEDFUNCTION()", error.Formula);
        Assert.Equal("#NAME?", RequireSuccess(_commands.GetValues(batch, sheetName, "A1")).Values[0][0]);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void RangeReads_WithStandardFormulaErrors_ReturnCanonicalNamesAndDiagnostics()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        string[] expectedNames = ["#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A"];
        int[] expectedCodes =
        [
            -2146826288,
            -2146826281,
            -2146826273,
            -2146826265,
            -2146826259,
            -2146826252,
            -2146826246
        ];

        RequireSuccess(_commands.SetFormulas(batch, sheetName, "A1:G1",
        [
            [
                "=SUM(A2:A3 C2:C3)",
                "=1/0",
                "=\"text\"+1",
                "=INDIRECT(\"A0\")",
                "=UNDEFINEDFUNCTION()",
                "=SQRT(-1)",
                "=NA()"
            ]
        ]));
        var calculation = _fixture.Send("calculationmode.calculate", new { scope = "application" });
        Assert.True(calculation.Success, calculation.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(calculation.ErrorMessage));

        var formulaResult = RequireSuccess(_commands.GetFormulas(batch, sheetName, "A1:G1"));
        var valueResult = RequireSuccess(_commands.GetValues(batch, sheetName, "A1:G1"));

        Assert.Equal(expectedNames, formulaResult.Values[0]);
        Assert.Equal(expectedNames, valueResult.Values[0]);

        Assert.NotNull(formulaResult.CellErrors);
        Assert.Equal(expectedNames.Length, formulaResult.CellErrors.Count);
        for (int index = 0; index < expectedNames.Length; index++)
        {
            var error = formulaResult.CellErrors[index];
            Assert.Equal($"{(char)('A' + index)}1", error.CellAddress);
            Assert.Equal(expectedCodes[index], error.ErrorCode);
            Assert.Equal(expectedCodes[index], error.CurrentValue);
            Assert.StartsWith(expectedNames[index], error.ErrorMessage, StringComparison.Ordinal);
            Assert.False(string.IsNullOrWhiteSpace(error.Suggestion));
        }

        string valueJson = System.Text.Json.JsonSerializer.Serialize(
            valueResult,
            System.Text.Json.JsonSerializerOptions.Web);
        using var valueDocument = System.Text.Json.JsonDocument.Parse(valueJson);
        var valueErrors = valueDocument.RootElement.GetProperty("cellErrors");
        Assert.Equal(expectedNames.Length, valueErrors.GetArrayLength());
        for (int index = 0; index < expectedNames.Length; index++)
        {
            var error = valueErrors[index];
            Assert.Equal($"{(char)('A' + index)}1", error.GetProperty("cellAddress").GetString());
            Assert.Equal(expectedNames[index], error.GetProperty("errorName").GetString());
            Assert.Equal(expectedCodes[index], error.GetProperty("errorCode").GetInt32());
            Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("formula").GetString()));
            Assert.False(string.IsNullOrWhiteSpace(error.GetProperty("suggestion").GetString()));
        }
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void GetFormulas_WithCircularReference_ReturnsExactFormulas()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Create circular reference: A1 = B1, B1 = A1
        RequireSuccess(_commands.SetFormulas(batch, sheetName, "A1", [
            ["=B1"]
        ]));
        RequireSuccess(_commands.SetFormulas(batch, sheetName, "B1", [
            ["=A1"]
        ]));

        // Act
        var result = _commands.GetFormulas(batch, sheetName, "A1:B1");

        // GetFormulas returns formula text; circular-reference diagnosis is not part of its contract.
        RequireSuccess(result);
        Assert.Equal(1, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal("=B1", result.Formulas[0][0]);
        Assert.Equal("=A1", result.Formulas[0][1]);
    }

    [Fact]
    [Trait("Feature", "Range")]
    public void GetFormulas_WithComplexRange_ReturnsAllErrorsWithAddresses()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set mix of valid and invalid formulas
        RequireSuccess(_commands.SetFormulas(batch, sheetName, "A1:A3", [
            ["=1+1"],                        // Valid
            ["=BADFUNCTION()"],              // Error
            ["=2+2"]                         // Valid
        ]));

        // Act
        var result = _commands.GetFormulas(batch, sheetName, "A1:A3");

        // Assert
        RequireSuccess(result);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Equal("=1+1", result.Formulas[0][0]);
        Assert.Equal("=BADFUNCTION()", result.Formulas[1][0]);
        Assert.Equal("=2+2", result.Formulas[2][0]);
        var values = RequireSuccess(_commands.GetValues(batch, sheetName, "A1:A3"));
        Assert.Equal(2, values.Values[0][0]);
        Assert.Equal("#NAME?", values.Values[1][0]);
        Assert.Equal(4, values.Values[2][0]);
        var error = Assert.Single(result.CellErrors!);
        Assert.Equal("A2", error.CellAddress);
        Assert.Equal("=BADFUNCTION()", error.Formula);
        Assert.Equal("#NAME?", error.ErrorName);
    }

    private void AssertCellUnchanged(
        IExcelBatch batch,
        string sheetName,
        string cellAddress,
        object? expectedValue,
        string? expectedFormula)
    {
        var values = RequireSuccess(_commands.GetValues(batch, sheetName, cellAddress));
        var formulas = RequireSuccess(_commands.GetFormulas(batch, sheetName, cellAddress));
        Assert.Equal(expectedValue, values.Values[0][0]);
        Assert.Equal(expectedFormula ?? "", formulas.Formulas[0][0]);
    }
}
