using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

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
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [[1234.56]]));

        // Apply LCID-based currency format
        RequireSuccess(_commands.SetNumberFormat(batch, sheetName, "A1", FormatCurrencyLCID));

        // Act - Read the displayed text and stored format directly from Excel
        var (displayedText, storedFormat, storedFormatLocal, rawValue) = ReadNumberDisplay(sheetName);
        var expectedText = ExpectedCurrencyDisplay();

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
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [[1234.56]]));

        // Apply format using NumberFormatLocal directly (locale-specific separators)
        // In German locale: , is decimal separator, . is thousands separator
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                cell = sheet.Range["A1"];
                cell.NumberFormatLocal =
                    $"$#{ctx.FormatTranslator.ThousandsSeparator}##0{ctx.FormatTranslator.DecimalSeparator}00";
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });

        // Act - Read the displayed text
        var (displayedText, storedFormat, storedFormatLocal, rawValue) = ReadNumberDisplay(sheetName);

        // Diagnostics
        _output.WriteLine($"Format applied (NumberFormatLocal): $#.##0,00");
        _output.WriteLine($"Format stored (NumberFormat): {storedFormat}");
        _output.WriteLine($"Format stored (NumberFormatLocal): {storedFormatLocal}");
        _output.WriteLine($"Displayed text: '{displayedText}'");

        // Assert
        Assert.Equal(ExpectedCurrencyDisplay(), displayedText);
        Assert.Equal(1234.56, rawValue);
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
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [[0.25]]));
        RequireSuccess(_commands.SetNumberFormat(batch, sheetName, "A1", FormatPercentage));

        // Act - Read the displayed text
        var (displayedText, _, _, rawValue) = ReadNumberDisplay(sheetName);

        // Assert
        _output.WriteLine($"Value: 0.25, Format: {FormatPercentage}");
        _output.WriteLine($"Displayed text: '{displayedText}'");

        Assert.Equal($"25{ReadDecimalSeparator()}00%", displayedText);
        Assert.Equal(0.25, rawValue);
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
        RequireSuccess(_commands.SetValues(batch, sheetName, "A1", [[1234567.89]]));
        RequireSuccess(_commands.SetNumberFormat(batch, sheetName, "A1", FormatNumber));

        // Act - Read the displayed text
        var (displayedText, _, _, rawValue) = ReadNumberDisplay(sheetName);

        // Assert
        _output.WriteLine($"Value: 1234567.89, Format: {FormatNumber}");
        _output.WriteLine($"Displayed text: '{displayedText}'");

        var thousands = _fixture.ExecuteRawVerification((ctx, _) => ctx.FormatTranslator.ThousandsSeparator);
        Assert.Equal($"1{thousands}234{thousands}567{ReadDecimalSeparator()}89", displayedText);
        Assert.Equal(1234567.89, rawValue);
    }

    private string ExpectedCurrencyDisplay() =>
        _fixture.ExecuteRawVerification((ctx, _) =>
            $"$1{ctx.FormatTranslator.ThousandsSeparator}234{ctx.FormatTranslator.DecimalSeparator}56");

    private (string Text, string Format, string LocalFormat, double Value) ReadNumberDisplay(string sheetName) =>
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? cell = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                cell = sheet.Range["A1"];
                string text = Convert.ToString(cell.Text, CultureInfo.InvariantCulture) ?? string.Empty;
                string format = Convert.ToString(cell.NumberFormat, CultureInfo.InvariantCulture) ?? string.Empty;
                string localFormat = Convert.ToString(cell.NumberFormatLocal, CultureInfo.InvariantCulture) ?? string.Empty;
                double value = Convert.ToDouble(cell.Value2, CultureInfo.InvariantCulture);
                return (text, format, localFormat, value);
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
            }
        });
}
