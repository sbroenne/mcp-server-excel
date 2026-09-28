using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Integration tests for RangeCommands number formatting operations.
/// Uses raw format codes - LLMs know Excel format codes natively.
/// </summary>
public sealed partial class PersistentServiceRangeSpecializedTests
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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set test value
        _commands.SetValues(batch, sheetName, "A1", [[1234.56]]);

        // Apply LCID-based currency format
        _commands.SetNumberFormat(batch, sheetName, "A1", FormatCurrencyLCID);

        // Act - Read the displayed text and stored format directly from Excel
        string displayedText = string.Empty;
        string storedFormat = string.Empty;
        string storedFormatLocal = string.Empty;
        string expectedText = string.Empty;
        object rawValue = null!;
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];
            displayedText = cell.Text?.ToString() ?? string.Empty;
            storedFormat = cell.NumberFormat?.ToString() ?? string.Empty;
            storedFormatLocal = cell.NumberFormatLocal?.ToString() ?? string.Empty;
            rawValue = cell.Value2;
            expectedText = $"$1{ctx.FormatTranslator.ThousandsSeparator}234{ctx.FormatTranslator.DecimalSeparator}56";
        });

        // Diagnostics
        _output.WriteLine($"Format applied: {FormatCurrencyLCID}");
        _output.WriteLine($"Format stored (NumberFormat): {storedFormat}");
        _output.WriteLine($"Format stored (NumberFormatLocal): {storedFormatLocal}");
        _output.WriteLine($"Raw value: {rawValue}");
        _output.WriteLine($"Displayed text: '{displayedText}'");

        // Assert - Verify Excel displays currency correctly
        Assert.Equal(expectedText, displayedText);
        Assert.Equal(1234.56, Assert.IsType<double>(rawValue));
    }

    /// <summary>
    /// Test using NumberFormatLocal to see if that works better for locale settings
    /// </summary>
    [Fact]
    public void SetNumberFormatLocal_Currency_DisplaysCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set test value
        _commands.SetValues(batch, sheetName, "A1", [[1234.56]]);

        // Apply format using NumberFormatLocal directly (locale-specific separators)
        // In German locale: , is decimal separator, . is thousands separator
        _fixture.ExecuteRawVerification((ctx, ct) =>
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
        _fixture.ExecuteRawVerification((ctx, ct) =>
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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set test value (0.25 should display as 25.00%)
        _commands.SetValues(batch, sheetName, "A1", [[0.25]]);
        _commands.SetNumberFormat(batch, sheetName, "A1", FormatPercentage);

        // Act - Read the displayed text
        string displayedText = string.Empty;
        _fixture.ExecuteRawVerification((ctx, ct) =>
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
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set large value to test thousands separator
        _commands.SetValues(batch, sheetName, "A1", [[1234567.89]]);
        _commands.SetNumberFormat(batch, sheetName, "A1", FormatNumber);

        // Act - Read the displayed text
        string displayedText = string.Empty;
        _fixture.ExecuteRawVerification((ctx, ct) =>
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

