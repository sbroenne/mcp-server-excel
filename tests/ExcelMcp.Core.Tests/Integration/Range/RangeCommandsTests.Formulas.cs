using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Range;

/// <summary>
/// Tests for range formulas operations
/// </summary>
public partial class RangeCommandsTests
{
    // === FORMULA OPERATIONS TESTS ===

    [Fact]
    public async Task GetFormulasAsync_ReturnsFormulasAndValues()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set values and formulas
        await _commands.SetValuesAsync(batch, "Sheet1", "A1:A3",
        [
            new() { 10 },
            new() { 20 },
            new() { 30 }
        ]);

        await _commands.SetFormulasAsync(batch, "Sheet1", "B1",
        [
            new() { "=SUM(A1:A3)" }
        ]);

        // Act
        var result = await _commands.GetFormulasAsync(batch, "Sheet1", "B1");

        // Assert
        Assert.True(result.Success);
        Assert.Equal("=SUM(A1:A3)", result.Formulas[0][0]);
        Assert.Equal(60.0, Convert.ToDouble(result.Values[0][0]));
    }

    [Fact]
    public async Task SetFormulasAsync_WritesFormulasToRange()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:A3",
        [
            new() { 5 },
            new() { 10 },
            new() { 15 }
        ]);

        var formulas = new List<List<string>>
        {
            new() { "=A1*2", "=A2*2", "=A3*2" }
        };

        // Act
        var result = await _commands.SetFormulasAsync(batch, "Sheet1", "B1:D1", formulas);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);

        // Verify values
        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "B1:D1");
        Assert.Equal(10.0, Convert.ToDouble(readResult.Values[0][0]));
        Assert.Equal(20.0, Convert.ToDouble(readResult.Values[0][1]));
        Assert.Equal(30.0, Convert.ToDouble(readResult.Values[0][2]));
    }

    [Fact]
    public async Task SetFormulasAsync_WithJsonElementFormulas_WritesFormulasCorrectly()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set up source data
        await _commands.SetValuesAsync(batch, "Sheet1", "A1:A3",
        [
            new() { 100 },
            new() { 200 },
            new() { 300 }
        ]);

        // Simulate MCP framework JSON deserialization
        // MCP receives: {"formulas": [["=SUM(A1:A3)", "=AVERAGE(A1:A3)"]]}
        // Framework deserializes to List<List<string>> where each string is JsonElement
        string json = """[["=SUM(A1:A3)", "=AVERAGE(A1:A3)"]]""";
        var jsonDoc = System.Text.Json.JsonDocument.Parse(json);

        var testFormulas = new List<List<string>>();
        foreach (var rowElement in jsonDoc.RootElement.EnumerateArray())
        {
            var row = new List<string>();
            foreach (var cellElement in rowElement.EnumerateArray())
            {
                // This is JsonElement, not primitive string
                row.Add(cellElement.GetString() ?? "");
            }
            testFormulas.Add(row);
        }

        // Act - Should handle JsonElement conversion internally
        var result = await _commands.SetFormulasAsync(batch, "Sheet1", "B1:C1", testFormulas);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"SetFormulasAsync failed: {result.ErrorMessage}");

        // Verify formulas were written correctly
        var formulaResult = await _commands.GetFormulasAsync(batch, "Sheet1", "B1:C1");
        Assert.True(formulaResult.Success);
        Assert.Equal("=SUM(A1:A3)", formulaResult.Formulas[0][0]);
        Assert.Equal("=AVERAGE(A1:A3)", formulaResult.Formulas[0][1]);

        // Verify calculated values
        Assert.Equal(600.0, Convert.ToDouble(formulaResult.Values[0][0])); // SUM
        Assert.Equal(200.0, Convert.ToDouble(formulaResult.Values[0][1])); // AVERAGE
    }
}
