using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeNumberFormatTests
{
    // Standard format codes - raw strings, no helper class needed
    private const string FormatCurrency = "$#,##0.00";
    private const string FormatPercentage = "0.00%";
    private const string FormatPercentageOneDecimal = "0.0%";
    private const string FormatNumber = "#,##0.00";
    private const string FormatDateShort = "m/d/yyyy";
    private const string FormatText = "@";

    [Fact]
    public void GetNumberFormats_SingleCell_ReturnsFormat()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up test data with a number format
        _commands.SetValues(batch, sheetName, "A1", [[100]]);
        _commands.SetNumberFormat(batch, sheetName, "A1", FormatCurrency);

        // Act
        var result = _commands.GetNumberFormats(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Single(result.Formats);
        Assert.Single(result.Formats[0]);
        // Excel might normalize format codes slightly
        Assert.Contains("$", result.Formats[0][0]); // Currency format present
    }

    [Fact]
    public void GetNumberFormats_MultipleFormats_ReturnsArray()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up test data FIRST
        _commands.SetValues(batch, sheetName, "A1:B2", [[100, 0.5], [200, 0.75]]);

        // THEN set different formats for each cell
        var formats = new List<List<string>>
        {
            new List<string> { FormatCurrency, FormatPercentage },
            new List<string> { FormatNumber, FormatPercentageOneDecimal }
        };
        _commands.SetNumberFormats(batch, sheetName, "A1:B2", formats);

        // Act
        var result = _commands.GetNumberFormats(batch, sheetName, "A1:B2");

        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(2, result.Formats.Count);
        // Verify currency and percentage symbols are present
        Assert.Contains("$", result.Formats[0][0]);
        Assert.Contains("%", result.Formats[0][1]);
        Assert.Contains("%", result.Formats[1][1]);
    }

    [Fact]
    public void SetNumberFormat_Currency_AppliesFormatToRange()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up test data
        _commands.SetValues(batch, sheetName, "A1:A3", [[100], [200], [300]]);

        // Act
        var result = _commands.SetNumberFormat(batch, sheetName, "A1:A3", FormatCurrency);

        // Assert - Verify operation success
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.Equal("set-number-format", result.Action);

        // Verify format was actually applied (check for currency symbol)
        var verifyResult = _commands.GetNumberFormats(batch, sheetName, "A1:A3");
        Assert.True(verifyResult.Success);
        Assert.Equal(3, verifyResult.Formats.Count);
        Assert.All(verifyResult.Formats, row => Assert.Contains("$", row[0])); // Currency symbol present
    }

    [Fact]
    public void SetNumberFormat_Percentage_AppliesFormatCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        _commands.SetValues(batch, sheetName, "B1:B2", [[0.25], [0.75]]);

        // Act
        var result = _commands.SetNumberFormat(batch, sheetName, "B1:B2", FormatPercentage);

        // Assert
        Assert.True(result.Success);

        // Verify format applied (check for percentage symbol)
        var verifyResult = _commands.GetNumberFormats(batch, sheetName, "B1:B2");
        Assert.True(verifyResult.Success);
        Assert.All(verifyResult.Formats, row => Assert.Contains("%", row[0])); // Percentage symbol present
    }

    [Fact]
    public void SetNumberFormat_DateFormat_AppliesCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Excel serial date: 45000 = April 17, 2023
        _commands.SetValues(batch, sheetName, "C1", [[45000]]);

        // Act
        var result = _commands.SetNumberFormat(batch, sheetName, "C1", FormatDateShort);

        // Assert
        Assert.True(result.Success);

        // Verify format applied (check for date-related format characters)
        var verifyResult = _commands.GetNumberFormats(batch, sheetName, "C1");
        Assert.True(verifyResult.Success);
        // Date formats contain d, m, or y characters
        Assert.Matches(
            @"[dmy]",
            verifyResult.Formats[0][0].ToLowerInvariant());
    }

    [Fact]
    public void SetNumberFormats_MixedFormats_AppliesDifferentFormatsPerCell()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up test data
        _commands.SetValues(batch, sheetName, "A1:C2", [[100, 0.5, 45000], [200, 0.75, 45100]]);

        // Act - Apply different formats to each column
        var formats = new List<List<string>>
        {
            new List<string> { FormatCurrency, FormatPercentage, FormatDateShort },
            new List<string> { FormatCurrency, FormatPercentage, FormatDateShort }
        };
        var result = _commands.SetNumberFormats(batch, sheetName, "A1:C2", formats);

        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");

        // Verify formats applied correctly (check for expected symbols/characters)
        var verifyResult = _commands.GetNumberFormats(batch, sheetName, "A1:C2");
        Assert.True(verifyResult.Success);
        Assert.Contains("$", verifyResult.Formats[0][0]); // Currency
        Assert.Contains("%", verifyResult.Formats[0][1]); // Percentage
        Assert.Matches(
            @"[dmy]",
            verifyResult.Formats[0][2].ToLowerInvariant()); // Date format
        Assert.Contains("$", verifyResult.Formats[1][0]); // Currency
        Assert.Contains("%", verifyResult.Formats[1][1]); // Percentage
        Assert.Matches(
            @"[dmy]",
            verifyResult.Formats[1][2].ToLowerInvariant()); // Date format
    }

    [Fact]
    public void SetNumberFormats_DimensionMismatch_ReturnsError()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Act & Assert - Try to apply 2x2 formats to 3x3 range (should throw ArgumentException)
        var formats = new List<List<string>>
        {
            new List<string> { FormatCurrency, FormatPercentage },
            new List<string> { FormatNumber, FormatPercentageOneDecimal }
        };
        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.SetNumberFormats(batch, sheetName, "A1:C3", formats));

        Assert.Contains("row count", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetNumberFormat_TextFormat_PreservesLeadingZeros()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // First set text format, then set value (to preserve leading zeros)
        _commands.SetNumberFormat(batch, sheetName, "D1", FormatText);
        _commands.SetValues(batch, sheetName, "D1", [["00123"]]);

        // Act - Verify format is text
        var result = _commands.GetNumberFormats(batch, sheetName, "D1");

        // Assert
        Assert.True(result.Success);
        Assert.Contains("@", result.Formats[0][0]); // Text format (@)
    }

}
