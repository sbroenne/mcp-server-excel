using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests for range values operations
/// </summary>
public partial class RangeCommandsTests
{
    // === VALUE OPERATIONS TESTS ===

    [Fact]
    public void GetValues_SingleCell_Returns1x1Array()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set a value first
        _commands.SetValues(batch, sheetName, "A1", [[100]]);

        // Act
        var result = _commands.GetValues(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Single(result.Values);
        Assert.Single(result.Values[0]);
        Assert.Equal(
            100.0,
            Convert.ToDouble(result.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void GetValues_3x3Range_Returns2DArray()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var testData = new List<List<object?>>
        {
            new() { 1, 2, 3 },
            new() { 4, 5, 6 },
            new() { 7, 8, 9 }
        };

        _commands.SetValues(batch, sheetName, "A1:C3", testData);

        // Act
        var result = _commands.GetValues(batch, sheetName, "A1:C3");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(3, result.Values.Count);
        Assert.Equal(
            1.0,
            Convert.ToDouble(result.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(
            9.0,
            Convert.ToDouble(result.Values[2][2], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void SetValues_TableWithHeaders_WritesAndReadsBack()
    {
        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var testData = new List<List<object?>>
        {
            new() { "Name", "Age" },
            new() { "Alice", 30 },
            new() { "Bob", 25 }
        };

        // Act
        var result = _commands.SetValues(batch, sheetName, "A1:B3", testData);
        // Assert
        Assert.True(result.Success);

        // Verify by reading back
        var readResult = _commands.GetValues(batch, sheetName, "A1:B3");
        Assert.Equal("Name", readResult.Values[0][0]);
        Assert.Equal(
            30.0,
            Convert.ToDouble(readResult.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void SetValues_JsonElementStrings_WritesCorrectly()
    {
        // Arrange - Simulate MCP Server scenario where JSON deserialization creates JsonElement objects
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Simulate MCP JSON: [["Azure Region Code", "Azure Region Name", "Geography", "Country"]]
        string json = """[["Azure Region Code", "Azure Region Name", "Geography", "Country"]]""";
        var jsonDoc = System.Text.Json.JsonDocument.Parse(json);
        var jsonArray = jsonDoc.RootElement;

        // Convert to List<List<object?>> containing JsonElement objects (like MCP does)
        var testData = new List<List<object?>>();
        foreach (var rowElement in jsonArray.EnumerateArray())
        {
            var row = new List<object?>();
            foreach (var cellElement in rowElement.EnumerateArray())
            {
                row.Add(cellElement); // This is a JsonElement, not a string!
            }
            testData.Add(row);
        }

        // Act
        var result = _commands.SetValues(batch, sheetName, "A1:D1", testData);
        // Assert
        Assert.True(result.Success, $"SetValuesAsync failed: {result.ErrorMessage}");

        // Verify by reading back
        var readResult = _commands.GetValues(batch, sheetName, "A1:D1");
        Assert.Equal("Azure Region Code", readResult.Values[0][0]);
        Assert.Equal("Azure Region Name", readResult.Values[0][1]);
        Assert.Equal("Geography", readResult.Values[0][2]);
        Assert.Equal("Country", readResult.Values[0][3]);
    }

    [Fact]
    public void SetValues_JsonElementMixedTypes_WritesCorrectly()
    {
        // Arrange - Test different JSON value types (string, number, boolean, null)
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Simulate MCP JSON: [["Text", 123, true, null]]
        string json = """[["Text", 123, true, null]]""";
        var jsonDoc = System.Text.Json.JsonDocument.Parse(json);
        var jsonArray = jsonDoc.RootElement;

        // Convert to List<List<object?>> containing JsonElement objects
        var testData = new List<List<object?>>();
        foreach (var rowElement in jsonArray.EnumerateArray())
        {
            var row = new List<object?>();
            foreach (var cellElement in rowElement.EnumerateArray())
            {
                row.Add(cellElement); // JsonElement
            }
            testData.Add(row);
        }

        // Act
        var result = _commands.SetValues(batch, sheetName, "A1:D1", testData);
        // Assert
        Assert.True(result.Success, $"SetValuesAsync failed: {result.ErrorMessage}");

        // Verify by reading back
        var readResult = _commands.GetValues(batch, sheetName, "A1:D1");
        Assert.Equal("Text", readResult.Values[0][0]);
        Assert.Equal(
            123.0,
            Convert.ToDouble(readResult.Values[0][1], System.Globalization.CultureInfo.InvariantCulture)); // Excel stores as double
        Assert.Equal(true, readResult.Values[0][2]);
        // Excel COM returns null (not empty string) for empty cells
        Assert.True(readResult.Values[0][3] == null || readResult.Values[0][3]?.ToString() == string.Empty);
    }

    [Fact]
    public void SetValues_WideHorizontalRange_NoOutOfMemoryError()
    {
        // Regression test for bug where 0-based arrays caused "out of memory" error
        // Root cause: Excel COM requires 1-based arrays, we were passing 0-based C# arrays

        // Arrange - use shared file, create unique sheet for this test
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Create test data with 16 columns (matching user's A2:P2 scenario)
        var testData = new List<List<object?>>
        {
            new object?[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16 }.ToList()
        };

        // Act - Write 16 values to A1:P1 (single row, 16 columns)
        var result = _commands.SetValues(batch, sheetName, "A1:P1", testData);

        // Assert - Should succeed without "out of memory" error
        Assert.True(result.Success, $"SetValues failed: {result.ErrorMessage}");

        // Verify values were written correctly
        var readResult = _commands.GetValues(batch, sheetName, "A1:P1");
        Assert.True(readResult.Success);
        Assert.Single(readResult.Values); // One row
        Assert.Equal(16, readResult.Values[0].Count); // 16 columns

        // Verify first, middle, and last values
        Assert.Equal(1.0, Convert.ToDouble(readResult.Values[0][0], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(8.0, Convert.ToDouble(readResult.Values[0][7], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal(16.0, Convert.ToDouble(readResult.Values[0][15], System.Globalization.CultureInfo.InvariantCulture));
    }

}
