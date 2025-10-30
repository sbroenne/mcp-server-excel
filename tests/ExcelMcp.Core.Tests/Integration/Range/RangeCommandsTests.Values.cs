using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Range;

/// <summary>
/// Tests for range values operations
/// </summary>
public partial class RangeCommandsTests
{
    // === VALUE OPERATIONS TESTS ===

    [Fact]
    public async Task GetValuesAsync_SingleCell_Returns1x1Array()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        // Set a value first
        await _commands.SetValuesAsync(batch, "Sheet1", "A1", [new() { 100 }]);

        // Act
        var result = await _commands.GetValuesAsync(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success, $"Failed: {result.ErrorMessage}");
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Single(result.Values);
        Assert.Single(result.Values[0]);
        Assert.Equal(100.0, Convert.ToDouble(result.Values[0][0]));
    }

    [Fact]
    public async Task GetValuesAsync_Range_Returns2DArray()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var testData = new List<List<object?>>
        {
            new() { 1, 2, 3 },
            new() { 4, 5, 6 },
            new() { 7, 8, 9 }
        };

        await _commands.SetValuesAsync(batch, "Sheet1", "A1:C3", testData);

        // Act
        var result = await _commands.GetValuesAsync(batch, "Sheet1", "A1:C3");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(3, result.Values.Count);
        Assert.Equal(1.0, Convert.ToDouble(result.Values[0][0]));
        Assert.Equal(9.0, Convert.ToDouble(result.Values[2][2]));
    }

    [Fact]
    public async Task SetValuesAsync_WritesDataToRange()
    {
        // Arrange
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

        var testData = new List<List<object?>>
        {
            new() { "Name", "Age" },
            new() { "Alice", 30 },
            new() { "Bob", 25 }
        };

        // Act
        var result = await _commands.SetValuesAsync(batch, "Sheet1", "A1:B3", testData);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success);

        // Verify by reading back
        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "A1:B3");
        Assert.Equal("Name", readResult.Values[0][0]);
        Assert.Equal(30.0, Convert.ToDouble(readResult.Values[1][1]));
    }

    [Fact]
    public async Task SetValuesAsync_WithJsonElementValues_WritesDataCorrectly()
    {
        // Arrange - Simulate MCP Server scenario where JSON deserialization creates JsonElement objects
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

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
        var result = await _commands.SetValuesAsync(batch, "Sheet1", "A1:D1", testData);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"SetValuesAsync failed: {result.ErrorMessage}");

        // Verify by reading back
        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "A1:D1");
        Assert.Equal("Azure Region Code", readResult.Values[0][0]);
        Assert.Equal("Azure Region Name", readResult.Values[0][1]);
        Assert.Equal("Geography", readResult.Values[0][2]);
        Assert.Equal("Country", readResult.Values[0][3]);
    }

    [Fact]
    public async Task SetValuesAsync_WithJsonElementMixedTypes_WritesDataCorrectly()
    {
        // Arrange - Test different JSON value types (string, number, boolean, null)
        string testFile = CreateTestWorkbook();
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);

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
        var result = await _commands.SetValuesAsync(batch, "Sheet1", "A1:D1", testData);
        await batch.SaveAsync();

        // Assert
        Assert.True(result.Success, $"SetValuesAsync failed: {result.ErrorMessage}");

        // Verify by reading back
        var readResult = await _commands.GetValuesAsync(batch, "Sheet1", "A1:D1");
        Assert.Equal("Text", readResult.Values[0][0]);
        Assert.Equal(123.0, Convert.ToDouble(readResult.Values[0][1])); // Excel stores as double
        Assert.Equal(true, readResult.Values[0][2]);
        // Excel COM returns null (not empty string) for empty cells
        Assert.True(readResult.Values[0][3] == null || readResult.Values[0][3]?.ToString() == string.Empty);
    }

}
