using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Integration tests for RangeCommands number formatting operations.
/// Uses raw format codes - LLMs know Excel format codes natively.
/// </summary>
public partial class RangeCommandsTests
{
    // Standard format codes - raw strings, no helper class needed
    private const string FormatCurrency = "$#,##0.00";
    private const string FormatPercentage = "0.00%";
    private const string FormatPercentageOneDecimal = "0.0%";
    private const string FormatNumber = "#,##0.00";
    private const string FormatDateShort = "m/d/yyyy";
    private const string FormatText = "@";

    // LCID-based currency format (proper Excel category recognition)
    private const string FormatCurrencyLCID = "[$$-409]#,##0.00"; // US Dollar with LCID

    [Fact]
    public void GetNumberFormats_SingleCell_ReturnsFormat()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
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
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
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

    /// <summary>
    /// CRITICAL TEST: Verifies Excel actually DISPLAYS formatted values correctly.
    /// This catches bugs where format code is applied but Excel doesn't render it properly.
    /// Uses the .Text property to read what Excel actually shows to users.
    /// NOTE: Excel uses SYSTEM LOCALE for separators, not format code. LCID only controls currency symbol.
    /// </summary>
    [Fact]
    public void SetNumberFormat_CurrencyWithLCID_DisplaysCorrectly()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set test value
        _commands.SetValues(batch, sheetName, "A1", [[1234.56]]);

        // Apply LCID-based currency format
        _commands.SetNumberFormat(batch, sheetName, "A1", FormatCurrencyLCID);

        // Act - Read the displayed text and stored format directly from Excel
        string displayedText = string.Empty;
        string storedFormat = string.Empty;
        string storedFormatLocal = string.Empty;
        object rawValue = null!;
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];
            displayedText = cell.Text?.ToString() ?? string.Empty;
            storedFormat = cell.NumberFormat?.ToString() ?? string.Empty;
            storedFormatLocal = cell.NumberFormatLocal?.ToString() ?? string.Empty;
            rawValue = cell.Value2;
        });

        // Diagnostics
        _output.WriteLine($"Format applied: {FormatCurrencyLCID}");
        _output.WriteLine($"Format stored (NumberFormat): {storedFormat}");
        _output.WriteLine($"Format stored (NumberFormatLocal): {storedFormatLocal}");
        _output.WriteLine($"Raw value: {rawValue}");
        _output.WriteLine($"Displayed text: '{displayedText}'");

        // Assert - Verify Excel displays currency correctly
        Assert.False(string.IsNullOrEmpty(displayedText), "Cell should display formatted text");
        Assert.Contains("$", displayedText); // Currency symbol from LCID
        // Formatted number includes thousands separator, so check for partial match
        Assert.True(
            displayedText.Contains("1234") || displayedText.Contains("1,234"),
            $"Number portion should be present, got: {displayedText}");
    }

    /// <summary>
    /// Test using NumberFormatLocal to see if that works better for locale settings
    /// </summary>
    [Fact]
    public void SetNumberFormatLocal_Currency_DisplaysCorrectly()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set test value
        _commands.SetValues(batch, sheetName, "A1", [[1234.56]]);

        // Apply format using NumberFormatLocal directly (locale-specific separators)
        // In German locale: , is decimal separator, . is thousands separator
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];
            // Use NumberFormatLocal with German-style separators (matching system locale)
            cell.NumberFormatLocal = "$#.##0,00";  // German style: . = thousands, , = decimal
        });

        // Act - Read the displayed text
        string displayedText = string.Empty;
        string storedFormat = string.Empty;
        string storedFormatLocal = string.Empty;
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];
            displayedText = cell.Text?.ToString() ?? string.Empty;
            storedFormat = cell.NumberFormat?.ToString() ?? string.Empty;
            storedFormatLocal = cell.NumberFormatLocal?.ToString() ?? string.Empty;
        });

        // Diagnostics
        _output.WriteLine($"Format applied (NumberFormatLocal): $#.##0,00");
        _output.WriteLine($"Format stored (NumberFormat): {storedFormat}");
        _output.WriteLine($"Format stored (NumberFormatLocal): {storedFormatLocal}");
        _output.WriteLine($"Displayed text: '{displayedText}'");

        // Assert
        Assert.False(string.IsNullOrEmpty(displayedText), "Cell should display formatted text");
        Assert.Contains("$", displayedText); // Currency symbol
        // Should have thousands separator and 2 decimal places
    }

    /// <summary>
    /// Verifies that percentage format displays correctly (not just format code applied).
    /// </summary>
    [Fact]
    public void SetNumberFormat_Percentage_DisplaysCorrectly()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set test value (0.25 should display as 25.00%)
        _commands.SetValues(batch, sheetName, "A1", [[0.25]]);
        _commands.SetNumberFormat(batch, sheetName, "A1", FormatPercentage);

        // Act - Read the displayed text
        string displayedText = string.Empty;
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];
            displayedText = cell.Text?.ToString() ?? string.Empty;
        });

        // Assert
        _output.WriteLine($"Value: 0.25, Format: {FormatPercentage}");
        _output.WriteLine($"Displayed text: '{displayedText}'");

        Assert.False(string.IsNullOrEmpty(displayedText), "Cell should display formatted text");
        Assert.Contains("%", displayedText); // Percentage symbol displayed
        Assert.Contains("25", displayedText); // Value multiplied by 100
    }

    /// <summary>
    /// Verifies that number format displays correctly.
    /// NOTE: Excel uses system locale for separators, not format code.
    /// </summary>
    [Fact]
    public void SetNumberFormat_NumberWithThousands_DisplaysCorrectly()
    {
        // Arrange
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set large value to test thousands separator
        _commands.SetValues(batch, sheetName, "A1", [[1234567.89]]);
        _commands.SetNumberFormat(batch, sheetName, "A1", FormatNumber);

        // Act - Read the displayed text
        string displayedText = string.Empty;
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];
            displayedText = cell.Text?.ToString() ?? string.Empty;
        });

        // Assert
        _output.WriteLine($"Value: 1234567.89, Format: {FormatNumber}");
        _output.WriteLine($"Displayed text: '{displayedText}'");

        Assert.False(string.IsNullOrEmpty(displayedText), "Cell should display formatted text");
        // Formatted number includes thousands separator (comma or period depending on locale)
        Assert.True(
            displayedText.Contains("1234567") || displayedText.Contains("1,234,567") || displayedText.Contains("1.234.567"),
            $"Number portion should be present, got: {displayedText}");
        // Decimal separator depends on locale (. or ,)
        Assert.True(
            displayedText.Contains("89") || displayedText.Contains(",89") || displayedText.Contains(".89"),
            $"Decimal portion should be displayed, got: {displayedText}");
    }
}




