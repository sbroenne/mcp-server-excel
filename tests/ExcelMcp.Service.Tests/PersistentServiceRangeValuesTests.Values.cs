using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for range values operations
/// </summary>
public sealed partial class PersistentServiceRangeValuesTests
{
    // === VALUE OPERATIONS TESTS ===

    [Fact]
    public async Task GetValues_MissingSheet_ReturnsCategorizedNotFound()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1", [["Retained"]]).Success);
        var response = await _fixture.SendForFailureAsync(
            "range.get-values",
            new { sheetName = "MissingSheet", rangeAddress = "A1" });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal("NotFound", response.ErrorCategory);
        var retained = _commands.GetValues(_fixture.BatchToken, sheetName, "A1");
        Assert.True(retained.Success, retained.ErrorMessage);
        Assert.Equal("Retained", Assert.Single(Assert.Single(retained.Values)));
    }

    [Theory]
    [InlineData("Not an address")]
    [InlineData("#")]
    [InlineData("A1##")]
    [InlineData("[]")]
    [InlineData("ReferenceTable[]")]
    [InlineData("ReferenceTable[Name]suffix")]
    public async Task GetValues_InvalidAddress_ReturnsCategorizedInvalidInput(
        string rangeAddress)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        Assert.True(_commands.SetValues(_fixture.BatchToken, sheetName, "A1:B1", [["Retained", 123]]).Success);

        var response = await _fixture.SendForFailureAsync(
            "range.get-values",
            new { sheetName, rangeAddress });

        Assert.Equal("OperationFailureException", response.ExceptionType);
        Assert.Equal("InvalidInput", response.ErrorCategory);
        var retained = _commands.GetValues(_fixture.BatchToken, sheetName, "A1:B1");
        Assert.True(retained.Success, retained.ErrorMessage);
        var cells = Assert.Single(retained.Values);
        Assert.Equal(2, cells.Count);
        Assert.Equal("Retained", cells[0]);
        Assert.Equal(123d, Convert.ToDouble(cells[1], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void GetValues_SingleCell_Returns1x1Array()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set a value first
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[100]]).Success);

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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var testData = new List<List<object?>>
        {
            new() { 1, 2, 3 },
            new() { 4, 5, 6 },
            new() { 7, 8, 9 }
        };

        Assert.True(_commands.SetValues(batch, sheetName, "A1:C3", testData).Success);

        // Act
        var result = _commands.GetValues(batch, sheetName, "A1:C3");

        // Assert
        Assert.True(result.Success);
        Assert.Equal(3, result.RowCount);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(3, result.Values.Count);
        for (var row = 0; row < testData.Count; row++)
        {
            Assert.Equal(testData[row].Count, result.Values[row].Count);
            for (var column = 0; column < testData[row].Count; column++)
            {
                Assert.Equal(Convert.ToDouble(testData[row][column], System.Globalization.CultureInfo.InvariantCulture),
                    Convert.ToDouble(result.Values[row][column], System.Globalization.CultureInfo.InvariantCulture));
            }
        }
    }

    [Fact]
    public void SetValues_TableWithHeaders_WritesAndReadsBack()
    {
        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
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
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal(3, readResult.RowCount);
        Assert.Equal(2, readResult.ColumnCount);
        Assert.Equal(["Name", "Age"], readResult.Values[0].Select(value => value?.ToString()));
        Assert.Equal("Alice", readResult.Values[1][0]);
        Assert.Equal(30d, Convert.ToDouble(readResult.Values[1][1], System.Globalization.CultureInfo.InvariantCulture));
        Assert.Equal("Bob", readResult.Values[2][0]);
        Assert.Equal(25d, Convert.ToDouble(readResult.Values[2][1], System.Globalization.CultureInfo.InvariantCulture));
    }

    [Fact]
    public void SetValues_JsonElementStrings_WritesCorrectly()
    {
        // Arrange - Simulate MCP Server scenario where JSON deserialization creates JsonElement objects
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Simulate MCP JSON: [["Azure Region Code", "Azure Region Name", "Geography", "Country"]]
        string json = """[["Azure Region Code", "Azure Region Name", "Geography", "Country"]]""";
        using var jsonDoc = System.Text.Json.JsonDocument.Parse(json);
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
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal(1, readResult.RowCount);
        Assert.Equal(4, readResult.ColumnCount);
        Assert.Equal(4, Assert.Single(readResult.Values).Count);
        Assert.Equal("Azure Region Code", readResult.Values[0][0]);
        Assert.Equal("Azure Region Name", readResult.Values[0][1]);
        Assert.Equal("Geography", readResult.Values[0][2]);
        Assert.Equal("Country", readResult.Values[0][3]);
    }

    [Fact]
    public void SetValues_JsonElementMixedTypes_WritesCorrectly()
    {
        // Arrange - Test different JSON value types (string, number, boolean, null)
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Simulate MCP JSON: [["Text", 123, true, null]]
        string json = """[["Text", 123, true, null]]""";
        using var jsonDoc = System.Text.Json.JsonDocument.Parse(json);
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
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal(1, readResult.RowCount);
        Assert.Equal(4, readResult.ColumnCount);
        Assert.Equal(4, Assert.Single(readResult.Values).Count);
        Assert.Equal("Text", readResult.Values[0][0]);
        Assert.Equal(
            123.0,
            Convert.ToDouble(readResult.Values[0][1], System.Globalization.CultureInfo.InvariantCulture)); // Excel stores as double
        Assert.Equal(true, readResult.Values[0][2]);
        // Excel COM returns null (not empty string) for empty cells
        Assert.True(readResult.Values[0][3] == null || readResult.Values[0][3]?.ToString() == string.Empty);
    }

    [Fact]
    public void SetValues_StrictIsoDate_WritesNativeExcelDate()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.SetValues(
            batch,
            sheetName,
            "A1:A2",
            [["Date"], ["2025-01-15"]]).Success);
        Assert.True(_commands.SetNumberFormat(
            batch,
            sheetName,
            "A2",
            "m/d/yyyy").Success);

        var readResult = _commands.GetValues(batch, sheetName, "A2");
        Assert.True(readResult.Success, readResult.ErrorMessage);
        var serial = Convert.ToDouble(
            readResult.Values[0][0],
            System.Globalization.CultureInfo.InvariantCulture);

        Assert.Equal(new DateTime(2025, 1, 15), DateTime.FromOADate(serial));

        Assert.True(_commands.SetValues(
            batch,
            sheetName,
            "B1:B2",
            [["Text"], ["'2025-01-15"]]).Success);

        var textResult = _commands.GetValues(batch, sheetName, "B2");
        Assert.True(textResult.Success, textResult.ErrorMessage);
        Assert.Equal("2025-01-15", textResult.Values[0][0]);
    }

    [Fact]
    public void SetValues_WideHorizontalRange_NoOutOfMemoryError()
    {
        // Regression test for bug where 0-based arrays caused "out of memory" error
        // Root cause: Excel COM requires 1-based arrays, we were passing 0-based C# arrays

        // Arrange - use shared file, create unique sheet for this test
        var batch = _fixture.BatchToken;
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

        Assert.Equal(Enumerable.Range(1, 16).Select(value => (double)value),
            readResult.Values[0].Select(value => Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture)));
    }

    [Fact]
    public void SetValues_AfterSheetCreate_ToNonA1Range_RoundTripsAndLeavesA1Empty()
    {
        var batch = _fixture.BatchToken;

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

        _fixture.CreateNamedTestSheet(batch, sheetName);

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

        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var jaggedValues = new List<List<object?>>
        {
            new object?[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14 }.ToList(),
            new object?[] { 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27 }.ToList()
        };
        var original = Enumerable.Range(1, 3)
            .Select(row => Enumerable.Range(1, 14)
                .Select(column => (object?)$"original-{row}-{column}").ToList()).ToList();
        var seeded = _commands.SetValues(batch, sheetName, "A1:N3", original);
        Assert.True(seeded.Success, seeded.ErrorMessage);

        var exception = Assert.Throws<ArgumentException>(
            () => _commands.SetValues(batch, sheetName, "A1:N2", jaggedValues));

        Assert.Contains("row 2", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("column count (13)", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("range column count (14)", exception.Message, StringComparison.OrdinalIgnoreCase);
        var retained = _commands.GetValues(batch, sheetName, "A1:N3");
        Assert.True(retained.Success, retained.ErrorMessage);
        Assert.Equal(original.Count, retained.Values.Count);
        for (var row = 0; row < original.Count; row++)
        {
            Assert.Equal(original[row], retained.Values[row]);
        }
    }

    [Fact]
    public void SetValues_MergedNonAnchorCell_ThrowsAndPreservesExistingValue()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["Original"]]).Success);
        Assert.True(_commands.MergeCells(batch, sheetName, "A1:B1").Success);

        var exception = Assert.Throws<InvalidOperationException>(
            () => _commands.SetValues(batch, sheetName, "B1", [["Updated"]]));

        Assert.Contains("$A$1:$B$1", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("top-left", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("unmerge", exception.Message, StringComparison.OrdinalIgnoreCase);

        var readResult = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal("Original", readResult.Values[0][0]);
    }

    [Fact]
    public void SetValues_MergedTopLeftCell_WritesAndRetainsValue()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.MergeCells(batch, sheetName, "A1:B1").Success);

        var result = _commands.SetValues(batch, sheetName, "A1", [["Updated"]]);

        Assert.True(result.Success, result.ErrorMessage);
        var readResult = _commands.GetValues(batch, sheetName, "A1");
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal("Updated", readResult.Values[0][0]);
    }

    [Fact]
    public void SetValues_RangeIntersectingMergedCells_ThrowsBeforeWriting()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["Anchor"]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "C1", [["Outside"]]).Success);
        Assert.True(_commands.MergeCells(batch, sheetName, "A1:B1").Success);

        var exception = Assert.Throws<InvalidOperationException>(
            () => _commands.SetValues(
                batch,
                sheetName,
                "A1:C1",
                [["New anchor", "Discarded", "New outside"]]));

        Assert.Contains("$A$1:$B$1", exception.Message, StringComparison.OrdinalIgnoreCase);
        var retained = _commands.GetValues(batch, sheetName, "A1:C1");
        Assert.True(retained.Success, retained.ErrorMessage);
        Assert.Equal("Anchor", retained.Values[0][0]);
        Assert.Null(retained.Values[0][1]);
        Assert.Equal("Outside", retained.Values[0][2]);
    }

    [Fact]
    public void SetValues_FormulaInMergedNonAnchorCell_Throws()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.SetValues(batch, sheetName, "A1", [["Original formula anchor"]]).Success);
        Assert.True(_commands.MergeCells(batch, sheetName, "A1:B1").Success);

        var exception = Assert.Throws<InvalidOperationException>(
            () => _commands.SetValues(batch, sheetName, "B1", [["=1+1"]]));

        Assert.Contains("$A$1:$B$1", exception.Message, StringComparison.OrdinalIgnoreCase);
        var retained = _commands.GetValues(batch, sheetName, "A1:B1");
        Assert.True(retained.Success, retained.ErrorMessage);
        Assert.Equal("Original formula anchor", retained.Values[0][0]);
        Assert.Null(retained.Values[0][1]);
    }

    [Fact]
    public void SetValues_MixedFormulaAndConstants_KeepsNonFormulaCells()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _fixture.Send("range.set-values", new
        {
            sheetName,
            rangeAddress = "A1:G1",
            values = new object?[][] { [1, "text", "=1+1", null, "2026-10-01", true, "'=not a formula"] }
        });

        var read = _commands.GetValues(batch, sheetName, "A1:G1");
        Assert.True(read.Success, read.ErrorMessage);
        var values = Assert.Single(read.Values);
        Assert.Equal(1d, Convert.ToDouble(values[0], CultureInfo.InvariantCulture));
        Assert.Equal("text", values[1]);
        Assert.Equal(2d, Convert.ToDouble(values[2], CultureInfo.InvariantCulture));
        Assert.Null(values[3]);
        Assert.Equal(46296d, Convert.ToDouble(values[4], CultureInfo.InvariantCulture));
        Assert.Equal(true, values[5]);
        Assert.Equal("=not a formula", values[6]);

        var hasFormula = _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cells = null;
            try
            {
                sheet = ComUtilities.FindSheet(context.Book, sheetName);
                Assert.NotNull(sheet);
                cells = sheet.Range["A1:G1"];
                var flags = new List<bool>();
                for (int column = 1; column <= 7; column++)
                {
                    Excel.Range? cell = null;
                    try
                    {
                        cell = (Excel.Range)cells.Cells[1, column];
                        flags.Add(Convert.ToBoolean(cell.HasFormula, CultureInfo.InvariantCulture));
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                    }
                }
                return flags;
            }
            finally
            {
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref sheet);
            }
        });
        Assert.Equal([false, false, true, false, false, false, false], hasFormula);
    }

    [Fact]
    public void GetValues_InvalidRange_ReportsSheetAndAddress()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var exception = Assert.Throws<InvalidOperationException>(() =>
            _commands.GetValues(batch, sheetName, "NotARange"));

        Assert.Contains($"Sheet '{sheetName}' exists", exception.Message);
        Assert.Contains("range 'NotARange' is invalid", exception.Message);
        Assert.Contains("Verify the range address format", exception.Message);
    }

}
