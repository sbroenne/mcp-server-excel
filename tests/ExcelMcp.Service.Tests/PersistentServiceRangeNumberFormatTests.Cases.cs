using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

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
        Assert.True(_commands.SetValues(batch, sheetName, "A1", [[100]]).Success);
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "A1", FormatCurrency).Success);

        // Act
        var result = _commands.GetNumberFormats(batch, sheetName, "A1");

        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.Equal(sheetName, result.SheetName);
        Assert.Equal(1, result.RowCount);
        Assert.Equal(1, result.ColumnCount);
        Assert.Single(result.Formats);
        Assert.Single(result.Formats[0]);
        AssertFormatMatrix([[FormatCurrency]], result.Formats);
        AssertNumericValues(sheetName, "A1", [[100]]);
    }

    [Fact]
    public void GetNumberFormats_MultipleFormats_ReturnsArray()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up test data FIRST
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B2", [[100, 0.5], [200, 0.75]]).Success);

        // THEN set different formats for each cell
        var formats = new List<List<string>>
        {
            new List<string> { FormatCurrency, FormatPercentage },
            new List<string> { FormatNumber, FormatPercentageOneDecimal }
        };
        Assert.True(_commands.SetNumberFormats(batch, sheetName, "A1:B2", formats).Success);

        // Act
        var result = _commands.GetNumberFormats(batch, sheetName, "A1:B2");

        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.Equal(2, result.RowCount);
        Assert.Equal(2, result.ColumnCount);
        Assert.Equal(2, result.Formats.Count);
        AssertFormatMatrix(formats, result.Formats);
        AssertNumericValues(sheetName, "A1:B2", [[100, 0.5], [200, 0.75]]);
    }

    [Fact]
    public void SetNumberFormat_Currency_AppliesFormatToRange()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up test data
        Assert.True(_commands.SetValues(batch, sheetName, "A1:A3", [[100], [200], [300]]).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "B1", [["Untouched"]]).Success);
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "B1", FormatText).Success);

        // Act
        var result = _commands.SetNumberFormat(batch, sheetName, "A1:A3", FormatCurrency);

        // Assert - Verify operation success
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");
        Assert.Equal("set-number-format", result.Action);

        var verifyResult = _commands.GetNumberFormats(batch, sheetName, "A1:A3");
        Assert.True(verifyResult.Success);
        Assert.Equal(3, verifyResult.Formats.Count);
        AssertFormatMatrix([[FormatCurrency], [FormatCurrency], [FormatCurrency]], verifyResult.Formats);
        AssertNumericValues(sheetName, "A1:A3", [[100], [200], [300]]);
        var untouchedFormat = _commands.GetNumberFormats(batch, sheetName, "B1");
        Assert.True(untouchedFormat.Success, untouchedFormat.ErrorMessage);
        AssertFormatMatrix([[FormatText]], untouchedFormat.Formats);
        var untouchedValue = _commands.GetValues(batch, sheetName, "B1");
        Assert.True(untouchedValue.Success, untouchedValue.ErrorMessage);
        Assert.Equal("Untouched", Assert.Single(Assert.Single(untouchedValue.Values)));
    }

    [Fact]
    public void SetNumberFormat_Percentage_AppliesFormatCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        Assert.True(_commands.SetValues(batch, sheetName, "B1:B2", [[0.25], [0.75]]).Success);

        // Act
        var result = _commands.SetNumberFormat(batch, sheetName, "B1:B2", FormatPercentage);

        // Assert
        Assert.True(result.Success);

        var verifyResult = _commands.GetNumberFormats(batch, sheetName, "B1:B2");
        Assert.True(verifyResult.Success);
        AssertFormatMatrix([[FormatPercentage], [FormatPercentage]], verifyResult.Formats);
        AssertNumericValues(sheetName, "B1:B2", [[0.25], [0.75]]);
    }

    [Fact]
    public void SetNumberFormat_DateFormat_AppliesCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Excel serial date 45000 is March 15, 2023.
        Assert.True(_commands.SetValues(batch, sheetName, "C1", [[45000]]).Success);

        // Act
        var result = _commands.SetNumberFormat(batch, sheetName, "C1", FormatDateShort);

        // Assert
        Assert.True(result.Success);

        var verifyResult = _commands.GetNumberFormats(batch, sheetName, "C1");
        Assert.True(verifyResult.Success);
        AssertFormatMatrix([[FormatDateShort]], verifyResult.Formats);
        var value = _commands.GetValues(batch, sheetName, "C1");
        Assert.True(value.Success, value.ErrorMessage);
        Assert.Equal(new DateTime(2023, 3, 15), DateTime.FromOADate(
            Convert.ToDouble(Assert.Single(Assert.Single(value.Values)),
                System.Globalization.CultureInfo.InvariantCulture)));
        AssertNativeDateDisplay(sheetName, "C1", [45000]);
    }

    [Fact]
    public void SetNumberFormats_MixedFormats_AppliesDifferentFormatsPerCell()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set up test data
        Assert.True(_commands.SetValues(batch, sheetName, "A1:C2", [[100, 0.5, 45000], [200, 0.75, 45100]]).Success);

        // Act - Apply different formats to each column
        var formats = new List<List<string>>
        {
            new List<string> { FormatCurrency, FormatPercentage, FormatDateShort },
            new List<string> { FormatCurrency, FormatPercentage, FormatDateShort }
        };
        var result = _commands.SetNumberFormats(batch, sheetName, "A1:C2", formats);

        // Assert
        Assert.True(result.Success, $"Operation failed: {result.ErrorMessage}");

        var verifyResult = _commands.GetNumberFormats(batch, sheetName, "A1:C2");
        Assert.True(verifyResult.Success);
        AssertFormatMatrix(formats, verifyResult.Formats);
        AssertNativeDateDisplay(sheetName, "C1:C2", [45000, 45100]);
        AssertNumericValues(sheetName, "A1:C2", [[100, 0.5, 45000], [200, 0.75, 45100]]);
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
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "A1:C3", FormatNumber).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:C3", [[1, 2, 3], [4, 5, 6], [7, 8, 9]]).Success);
        var original = _commands.GetNumberFormats(batch, sheetName, "A1:C3");
        Assert.True(original.Success, original.ErrorMessage);
        Assert.All(original.Formats, row => Assert.All(row, format => Assert.Equal(FormatNumber, format)));
        var exception = Assert.Throws<ArgumentException>(() =>
            _commands.SetNumberFormats(batch, sheetName, "A1:C3", formats));

        Assert.Contains("row count", exception.Message, StringComparison.OrdinalIgnoreCase);
        var retained = _commands.GetNumberFormats(batch, sheetName, "A1:C3");
        Assert.True(retained.Success, retained.ErrorMessage);
        AssertFormatMatrix(original.Formats, retained.Formats);
        AssertNumericValues(sheetName, "A1:C3", [[1, 2, 3], [4, 5, 6], [7, 8, 9]]);
    }

    [Fact]
    public void SetNumberFormat_TextFormat_PreservesLeadingZeros()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // First set text format, then set value (to preserve leading zeros)
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "D1", FormatText).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "D1", [["00123"]]).Success);

        // Act - Verify format is text
        var result = _commands.GetNumberFormats(batch, sheetName, "D1");

        // Assert
        Assert.True(result.Success);
        AssertFormatMatrix([[FormatText]], result.Formats);
        var value = _commands.GetValues(batch, sheetName, "D1");
        Assert.True(value.Success, value.ErrorMessage);
        Assert.Equal("00123", Assert.Single(Assert.Single(value.Values)));
    }

    [Fact]
    public void SetNumberFormats_LaterRowMismatch_PreservesExistingFormats()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        Assert.True(_commands.SetNumberFormat(batch, sheetName, "A1:B2", FormatPercentageOneDecimal).Success);
        Assert.True(_commands.SetValues(batch, sheetName, "A1:B2", [[0.1, 0.2], [0.3, 0.4]]).Success);
        var original = _commands.GetNumberFormats(batch, sheetName, "A1:B2");
        Assert.True(original.Success, original.ErrorMessage);
        AssertFormatMatrix(
            [[FormatPercentageOneDecimal, FormatPercentageOneDecimal], [FormatPercentageOneDecimal, FormatPercentageOneDecimal]],
            original.Formats);

        var rejected = Assert.Throws<ArgumentException>(() =>
            _commands.SetNumberFormats(batch, sheetName, "A1:B2",
                [[FormatCurrency, FormatNumber], [FormatText]]));
        Assert.Contains("row 2 column count", rejected.Message, StringComparison.OrdinalIgnoreCase);
        var retained = _commands.GetNumberFormats(batch, sheetName, "A1:B2");
        Assert.True(retained.Success, retained.ErrorMessage);
        AssertFormatMatrix(original.Formats, retained.Formats);
        AssertNumericValues(sheetName, "A1:B2", [[0.1, 0.2], [0.3, 0.4]]);
    }

    private void AssertNumericValues(string sheetName, string address, List<List<double>> expected)
    {
        var result = _commands.GetValues(_fixture.BatchToken, sheetName, address);
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Equal(expected.Count, result.RowCount);
        Assert.Equal(expected[0].Count, result.ColumnCount);
        Assert.Equal(expected.Count, result.Values.Count);
        for (var row = 0; row < expected.Count; row++)
        {
            Assert.Equal(expected[row].Count, result.Values[row].Count);
            Assert.Equal(expected[row], result.Values[row].Select(value =>
                Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture)));
        }
    }

    private static void AssertFormatMatrix(List<List<string>> expected, List<List<string>> actual)
    {
        Assert.Equal(expected.Count, actual.Count);
        for (var row = 0; row < expected.Count; row++)
        {
            Assert.Equal(expected[row].Count, actual[row].Count);
            for (var column = 0; column < expected[row].Count; column++)
            {
                Assert.Equal(expected[row][column],
                    actual[row][column].Replace("\\$", "$", StringComparison.Ordinal)
                        .Replace("\\/", "/", StringComparison.Ordinal));
            }
        }

    }

    private void AssertNativeDateDisplay(string sheetName, string address, double[] serials) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Excel.Range? cells = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                range = sheet.Range[address];
                range.ColumnWidth = 20;
                cells = range.Cells;
                Assert.Equal(serials.Length, cells.Count);
                for (var index = 0; index < serials.Length; index++)
                {
                    Excel.Range? cell = null;
                    try
                    {
                        cell = (Excel.Range)cells[index + 1];
                        Assert.Equal(serials[index], Convert.ToDouble(cell.Value2,
                            System.Globalization.CultureInfo.InvariantCulture));
                        Assert.Equal(DateTime.FromOADate(serials[index]).ToString(
                            "M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture), cell.Text);
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref cells);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
}
