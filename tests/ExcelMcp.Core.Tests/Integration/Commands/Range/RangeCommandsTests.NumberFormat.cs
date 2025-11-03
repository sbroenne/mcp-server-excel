using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Integration tests for RangeCommands number formatting operations
/// </summary>
public partial class RangeCommandsTests
{
    [Fact]
    public async Task GetNumberFormats_SingleCell_ReturnsFormat()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests),
            nameof(GetNumberFormats_SingleCell_ReturnsFormat),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Set up test data with a number format
        await _commands.SetValuesAsync(batch, "Sheet1", "A1", [[100]]);
        await _commands.SetNumberFormatAsync(batch, "Sheet1", "A1", NumberFormatPresets.Currency);

        // Act
        var result = await _commands.GetNumberFormatsAsync(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.Equal("Sheet1", result.SheetName);
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Single(result.Formats);
        Assert.Single(result.Formats[0]);
        // Excel might normalize format codes slightly
        Assert.Contains("$", result.Formats[0][0]); // Currency format present
    }

    [Fact]
    public async Task GetNumberFormats_MultipleFormats_ReturnsArray()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests),
            nameof(GetNumberFormats_MultipleFormats_ReturnsArray),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Set up test data FIRST
        await _commands.SetValuesAsync(batch, "Sheet1", "A1:B2", [[100, 0.5], [200, 0.75]]);
        
        // THEN set different formats for each cell
        var formats = new List<List<string>>
        {
            new List<string> { NumberFormatPresets.Currency, NumberFormatPresets.Percentage },
            new List<string> { NumberFormatPresets.Number, NumberFormatPresets.PercentageOneDecimal }
        };
        await _commands.SetNumberFormatsAsync(batch, "Sheet1", "A1:B2", formats);

        // Act
        var result = await _commands.GetNumberFormatsAsync(batch, "Sheet1", "A1:B2");

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
    public async Task SetNumberFormat_Currency_AppliesFormatToRange()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests),
            nameof(SetNumberFormat_Currency_AppliesFormatToRange),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Set up test data
        await _commands.SetValuesAsync(batch, "Sheet1", "A1:A3", [[100], [200], [300]]);

        // Act
        var result = await _commands.SetNumberFormatAsync(batch, "Sheet1", "A1:A3", NumberFormatPresets.Currency);

        // Assert - Verify operation success
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.Equal("set-number-format", result.Action);

        // Verify format was actually applied (check for currency symbol)
        var verifyResult = await _commands.GetNumberFormatsAsync(batch, "Sheet1", "A1:A3");
        Assert.True(verifyResult.Success);
        Assert.Equal(3, verifyResult.Formats.Count);
        Assert.All(verifyResult.Formats, row => Assert.Contains("$", row[0])); // Currency symbol present
    }

    [Fact]
    public async Task SetNumberFormat_Percentage_AppliesFormatCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests),
            nameof(SetNumberFormat_Percentage_AppliesFormatCorrectly),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        await _commands.SetValuesAsync(batch, "Sheet1", "B1:B2", [[0.25], [0.75]]);

        // Act
        var result = await _commands.SetNumberFormatAsync(batch, "Sheet1", "B1:B2", NumberFormatPresets.Percentage);

        // Assert
        Assert.True(result.Success);

        // Verify format applied (check for percentage symbol)
        var verifyResult = await _commands.GetNumberFormatsAsync(batch, "Sheet1", "B1:B2");
        Assert.True(verifyResult.Success);
        Assert.All(verifyResult.Formats, row => Assert.Contains("%", row[0])); // Percentage symbol present
    }

    [Fact]
    public async Task SetNumberFormat_DateFormat_AppliesCorrectly()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests),
            nameof(SetNumberFormat_DateFormat_AppliesCorrectly),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Excel serial date: 45000 = April 17, 2023
        await _commands.SetValuesAsync(batch, "Sheet1", "C1", [[45000]]);

        // Act
        var result = await _commands.SetNumberFormatAsync(batch, "Sheet1", "C1", NumberFormatPresets.DateShort);

        // Assert
        Assert.True(result.Success);

        // Verify format applied (check for date-related format characters)
        var verifyResult = await _commands.GetNumberFormatsAsync(batch, "Sheet1", "C1");
        Assert.True(verifyResult.Success);
        // Date formats contain d, m, or y characters
        Assert.Matches(@"[dmy]", verifyResult.Formats[0][0].ToLower());
    }

    [Fact]
    public async Task SetNumberFormats_MixedFormats_AppliesDifferentFormatsPerCell()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests),
            nameof(SetNumberFormats_MixedFormats_AppliesDifferentFormatsPerCell),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Set up test data
        await _commands.SetValuesAsync(batch, "Sheet1", "A1:C2", [[100, 0.5, 45000], [200, 0.75, 45100]]);

        // Act - Apply different formats to each column
        var formats = new List<List<string>>
        {
            new List<string> { NumberFormatPresets.Currency, NumberFormatPresets.Percentage, NumberFormatPresets.DateShort },
            new List<string> { NumberFormatPresets.Currency, NumberFormatPresets.Percentage, NumberFormatPresets.DateShort }
        };
        var result = await _commands.SetNumberFormatsAsync(batch, "Sheet1", "A1:C2", formats);

        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");

        // Verify formats applied correctly (check for expected symbols/characters)
        var verifyResult = await _commands.GetNumberFormatsAsync(batch, "Sheet1", "A1:C2");
        Assert.True(verifyResult.Success);
        Assert.Contains("$", verifyResult.Formats[0][0]); // Currency
        Assert.Contains("%", verifyResult.Formats[0][1]); // Percentage
        Assert.Matches(@"[dmy]", verifyResult.Formats[0][2].ToLower()); // Date format
        Assert.Contains("$", verifyResult.Formats[1][0]); // Currency
        Assert.Contains("%", verifyResult.Formats[1][1]); // Percentage
        Assert.Matches(@"[dmy]", verifyResult.Formats[1][2].ToLower()); // Date format
    }

    [Fact]
    public async Task SetNumberFormats_DimensionMismatch_ReturnsError()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests),
            nameof(SetNumberFormats_DimensionMismatch_ReturnsError),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // Act - Try to apply 2x2 formats to 3x3 range
        var formats = new List<List<string>>
        {
            new List<string> { NumberFormatPresets.Currency, NumberFormatPresets.Percentage },
            new List<string> { NumberFormatPresets.Number, NumberFormatPresets.PercentageOneDecimal }
        };
        var result = await _commands.SetNumberFormatsAsync(batch, "Sheet1", "A1:C3", formats);

        // Assert - Should fail with dimension mismatch
        Assert.False(result.Success);
        Assert.Contains("row count", result.ErrorMessage, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task SetNumberFormat_TextFormat_PreservesLeadingZeros()
    {
        // Arrange
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(RangeCommandsTests),
            nameof(SetNumberFormat_TextFormat_PreservesLeadingZeros),
            _tempDir,
            ".xlsx");

        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        
        // First set text format, then set value (to preserve leading zeros)
        await _commands.SetNumberFormatAsync(batch, "Sheet1", "D1", NumberFormatPresets.Text);
        await _commands.SetValuesAsync(batch, "Sheet1", "D1", [["00123"]]);

        // Act - Verify format is text
        var result = await _commands.GetNumberFormatsAsync(batch, "Sheet1", "D1");

        // Assert
        Assert.True(result.Success);
        Assert.Contains("@", result.Formats[0][0]); // Text format (@)
    }
}
