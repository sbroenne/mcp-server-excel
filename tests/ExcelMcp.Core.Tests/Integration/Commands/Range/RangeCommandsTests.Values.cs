using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
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

    [Fact]
    public void SetValues_AfterSheetCreate_ToNonA1Range_RoundTripsAndLeavesA1Empty()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);

        var sheetCommands = new SheetCommands();
        var sheetName = $"Bug2_{Guid.NewGuid():N}"[..31];
        var values = new List<List<object?>>
        {
            new() { "R1C1", "R1C2", "R1C3", "R1C4", "R1C5", "R1C6", "R1C7" },
            new() { "R2C1", "R2C2", "R2C3", "R2C4", "R2C5", "R2C6", "R2C7" },
            new() { "R3C1", "R3C2", "R3C3", "R3C4", "R3C5", "R3C6", "R3C7" },
            new() { "R4C1", "R4C2", "R4C3", "R4C4", "R4C5", "R4C6", "R4C7" },
            new() { "R5C1", "R5C2", "R5C3", "R5C4", "R5C5", "R5C6", "R5C7" },
            new() { "R6C1", "R6C2", "R6C3", "R6C4", "R6C5", "R6C6", "R6C7" },
            new() { "R7C1", "R7C2", "R7C3", "R7C4", "R7C5", "R7C6", "R7C7" },
            new() { "R8C1", "R8C2", "R8C3", "R8C4", "R8C5", "R8C6", "R8C7" }
        };

        sheetCommands.Create(batch, sheetName);

        var writeResult = _commands.SetValues(batch, sheetName, "A3:G10", values);

        Assert.True(writeResult.Success, $"SetValues failed: {writeResult.ErrorMessage}");

        var readResult = _commands.GetValues(batch, sheetName, "A3:G10");
        Assert.True(readResult.Success, $"GetValues failed: {readResult.ErrorMessage}");
        Assert.Equal(8, readResult.RowCount);
        Assert.Equal(7, readResult.ColumnCount);

        for (int rowIndex = 0; rowIndex < values.Count; rowIndex++)
        {
            for (int columnIndex = 0; columnIndex < values[rowIndex].Count; columnIndex++)
            {
                Assert.Equal(values[rowIndex][columnIndex], readResult.Values[rowIndex][columnIndex]);
            }
        }

        var a1Result = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(a1Result.Success, $"GetValues A1 failed: {a1Result.ErrorMessage}");
        Assert.True(a1Result.Values[0][0] == null || a1Result.Values[0][0]?.ToString() == string.Empty);
    }

    [Fact]
    public void SetValues_JaggedWideRange_ThrowsDescriptiveValidationError()
    {
        // Regression test for Bug 1 root cause hypothesis:
        // wide writes only fail when later rows are shorter than the first row.

        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        var jaggedValues = new List<List<object?>>
        {
            new object?[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14 }.ToList(),
            new object?[] { 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27 }.ToList()
        };

        var exception = Assert.Throws<ArgumentException>(
            () => _commands.SetValues(batch, sheetName, "A1:N2", jaggedValues));

        Assert.Contains("row 2", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("column count (13)", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("range column count (14)", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

}




