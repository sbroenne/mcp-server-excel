using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests that format translation (date and number separators) works correctly across locales.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Service")]
[Trait("Feature", "Ranges")]
[Collection("ServiceWorkflow")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServiceFormatTranslationTests(
    PersistentServiceWorkbookFixture fixture,
    ITestOutputHelper output) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly ITestOutputHelper _output = output;
    private readonly IRangeServiceCommands _rangeCommands =
        fixture.CreateCommands<IRangeServiceCommands>();

    [Theory]
    [InlineData("mmm-yy")]
    [InlineData("dddd, mmmm d, yyyy")]
    [InlineData("General")]
    public void SetNumberFormat_LocalKeywordsAndDateNames_MatchNativeDisplay(string format)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1:B1", [[45000.75, 45000.75]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", format));
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.Range? actual = null;
            Microsoft.Office.Interop.Excel.Range? native = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[sheetName];
                actual = sheet.Range["A1"];
                native = sheet.Range["B1"];
                actual.ColumnWidth = native.ColumnWidth = 50;
                native.NumberFormat = format;
                _output.WriteLine($"Requested '{format}': native '{native.NumberFormat}' = '{native.Text}', service '{actual.NumberFormat}' = '{actual.Text}'");
                if (format == "mmm-yy")
                {
                    // Excel's built-in month/year format substitutes regional punctuation.
                    // Render each date component independently to check the requested hyphen.
                    native.NumberFormat = "mmm";
                    var month = Assert.IsType<string>(native.Text);
                    native.NumberFormat = "yy";
                    var year = Assert.IsType<string>(native.Text);
                    Assert.Equal($"{month}-{year}", actual.Text);
                }
                else
                {
                    Assert.Equal(native.Text, actual.Text);
                    Assert.Equal(native.NumberFormat, actual.NumberFormat);
                }
                Assert.Equal(format, actual.NumberFormat);
                Assert.Equal(45000.75, Assert.IsType<double>(actual.Value2));
                Assert.Equal(45000.75, Assert.IsType<double>(native.Value2));
            }
            finally
            {
                ComUtilities.Release(ref native);
                ComUtilities.Release(ref actual);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Theory]
    [InlineData("[$$-409]#,##0.00", 1234.56, "$1{group}234{decimal}56")]
    [InlineData("yyyy-mm-dd hh:mm:ss", 45000.75, "2023-03-15 18:00:00")]
    [InlineData("[h]:mm:ss", 1.75, "42:00:00")]
    [InlineData("0.00,,\"M\"", 1250000, "1{decimal}25M")]
    [InlineData("[>=1.5]\"high\";\"low\"", 1.25, "low")]
    [InlineData("[>=1.5]\"high\";\"low\"", 1.75, "high")]
    [InlineData("0.00 \"a,b.c\"\\!", 12.5, "12{decimal}50 a,b.c!")]
    public void NativeRangeNumberFormat_InvariantCodes_PreserveDisplayMeaning(
        string format, double value, string expectedText)
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.Range? cell = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[sheetName];
                cell = sheet.Range["A1"];
                cell.ColumnWidth = 40;
                cell.Value2 = value;
                cell.NumberFormat = format;
                _output.WriteLine($"Native NumberFormat: '{cell.NumberFormat}', local: '{cell.NumberFormatLocal}', text: '{cell.Text}'");
                Assert.Equal(expectedText
                    .Replace("{decimal}", ctx.FormatTranslator.DecimalSeparator, StringComparison.Ordinal)
                    .Replace("{group}", ctx.FormatTranslator.ThousandsSeparator, StringComparison.Ordinal),
                    cell.Text);
                Assert.Equal(value, Assert.IsType<double>(cell.Value2));
            }
            finally
            {
                ComUtilities.Release(ref cell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Fact]
    public void SetNumberFormat_DollarLiteral_PreservesCurrencyUnlikeNativeInvariantProperty()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1:B1", [[1234.56, 1234.56]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", "$#,##0.00"));
        var read = RequireSuccess(_rangeCommands.GetNumberFormats(batch, sheetName, "A1"));
        var returnedFormat = Assert.Single(Assert.Single(read.Formats));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", returnedFormat));
        var reread = RequireSuccess(_rangeCommands.GetNumberFormats(batch, sheetName, "A1"));
        Assert.Equal(returnedFormat, Assert.Single(Assert.Single(reread.Formats)));
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.Range? writtenCell = null;
            Microsoft.Office.Interop.Excel.Range? nativeCell = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[sheetName];
                writtenCell = sheet.Range["A1"];
                nativeCell = sheet.Range["B1"];
                Assert.Equal(returnedFormat, writtenCell.NumberFormat);
                nativeCell.NumberFormat = "$#,##0.00";
                var digits = $"1{ctx.FormatTranslator.ThousandsSeparator}234{ctx.FormatTranslator.DecimalSeparator}56";
                var nativeCurrency = Convert.ToString(
                    ctx.App.International[Microsoft.Office.Interop.Excel.XlApplicationInternational.xlCurrencyCode],
                    CultureInfo.InvariantCulture);
                Assert.Equal("$" + digits, writtenCell.Text);
                Assert.Equal(nativeCurrency + digits, nativeCell.Text);
                Assert.Equal(1234.56, Assert.IsType<double>(writtenCell.Value2));
                Assert.Equal(1234.56, Assert.IsType<double>(nativeCell.Value2));
            }
            finally
            {
                ComUtilities.Release(ref nativeCell);
                ComUtilities.Release(ref writtenCell);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Theory]
    [InlineData("single")]
    [InlineData("column")]
    [InlineData("grid")]
    [InlineData("ranges")]
    public void NumberFormatWrites_ComplexCodes_PreserveFormatsAndDisplayMeaning(string action)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        string[] formats =
        [
            "\\$#,##0.00", "yyyy-mm-dd hh:mm:ss", "[h]:mm:ss", "0.00,,\"M\"",
            "[>=1.5]\"high\";\"low\"", "[>=1.5]\"high\";\"low\"", "0.00 \"a,b.c\"\\m", "[$$-409]#,##0.00"
        ];
        double[] values = [1234.56, 45000.75, 1.75, 1250000, 1.25, 1.75, 12.5, 1234.56];
        string[] expected =
        [
            "$1{group}234{decimal}56", "2023-03-15 18:00:00", "42:00:00", "1{decimal}25M",
            "low", "high", "12{decimal}50 a,b.cm", "$1{group}234{decimal}56"
        ];
        var columns = action == "grid" ? 2 : 1;
        var rows = formats.Length / columns;
        var address = columns == 2 ? "A1:B4" : "A1:A8";
        var data = Enumerable.Range(0, rows)
            .Select(row => Enumerable.Range(0, columns).Select(col => (object?)values[row * columns + col]).ToList()).ToList();
        var codes = Enumerable.Range(0, rows)
            .Select(row => Enumerable.Range(0, columns).Select(col => formats[row * columns + col]).ToList()).ToList();
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, address, data));
        if (action is "column" or "grid")
        {
            var written = _rangeCommands.SetNumberFormats(batch, sheetName, address, codes);
            RequireSuccess(written);
        }
        else
        {
            for (var index = 0; index < formats.Length; index++)
            {
                var cellAddress = $"A{index + 1}";
                var written = action == "single"
                    ? _rangeCommands.SetNumberFormat(batch, sheetName, cellAddress, formats[index])
                    : _rangeCommands.Format(batch, sheetName, [cellAddress],
                        new() { NumberFormat = formats[index] });
                RequireSuccess(written);
            }
        }
        var read = RequireSuccess(_rangeCommands.GetNumberFormats(batch, sheetName, address));
        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.Range? range = null;
            Microsoft.Office.Interop.Excel.Range? cells = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[sheetName];
                range = sheet.Range[address];
                range.ColumnWidth = 40;
                cells = range.Cells;
                for (var row = 0; row < rows; row++)
                {
                    for (var col = 0; col < columns; col++)
                    {
                        Microsoft.Office.Interop.Excel.Range? cell = null;
                        try
                        {
                            var index = row * columns + col;
                            cell = (Microsoft.Office.Interop.Excel.Range)cells[row + 1, col + 1];
                            Assert.Equal(formats[index], read.Formats[row][col]);
                            Assert.Equal(expected[index]
                                .Replace("{decimal}", ctx.FormatTranslator.DecimalSeparator, StringComparison.Ordinal)
                                .Replace("{group}", ctx.FormatTranslator.ThousandsSeparator, StringComparison.Ordinal), cell.Text);
                            Assert.Equal(values[index], Assert.IsType<double>(cell.Value2));
                        }
                        finally
                        {
                            ComUtilities.Release(ref cell);
                        }
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

    [Theory]
    [InlineData("en-US")]
    [InlineData("de-DE")]
    public void NumberFormats_InvariantCodes_RoundTripAcrossCallerCultures(string cultureName)
    {
        var originalCulture = CultureInfo.CurrentCulture;
        try
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(cultureName);
            var batch = _fixture.BatchToken;
            var sheetName = _fixture.CreateTestSheet(batch);
            RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1", [[0.125]]));
            RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", "0.00%"));
            var result = RequireSuccess(_rangeCommands.GetNumberFormats(batch, sheetName, "A1"));
            Assert.Equal("0.00%", result.Formats[0][0]);
            _fixture.ExecuteRawVerification((ctx, _) =>
            {
                Microsoft.Office.Interop.Excel.Sheets? sheets = null;
                Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
                Microsoft.Office.Interop.Excel.Range? cell = null;
                try
                {
                    sheets = ctx.Book.Worksheets;
                    sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[sheetName];
                    cell = sheet.Range["A1"];
                    Assert.Equal(0.125, Assert.IsType<double>(cell.Value2));
                    Assert.Equal($"12{ctx.FormatTranslator.DecimalSeparator}50%", cell.Text);
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            });
        }
        finally
        {
            CultureInfo.CurrentCulture = originalCulture;
        }
    }

    [Fact]
    public void SetNumberFormat_DecimalCondition_PreservesTheComparisonThreshold()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1:A2", [[1.25], [1.75]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1:A2", "[>=1.5]\"high\";\"low\""));

        _fixture.ExecuteRawVerification((ctx, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.Range? below = null;
            Microsoft.Office.Interop.Excel.Range? above = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[sheetName];
                below = sheet.Range["A1"];
                above = sheet.Range["A2"];
                Assert.Equal("low", below.Text);
                Assert.Equal("high", above.Text);
                Assert.Equal(1.25, Assert.IsType<double>(below.Value2));
                Assert.Equal(1.75, Assert.IsType<double>(above.Value2));
            }
            finally
            {
                ComUtilities.Release(ref above);
                ComUtilities.Release(ref below);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
    }

    [Fact]
    public void SetNumberFormat_USDateFormat_DisplaysCorrectly()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1", [[45000]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", "m/d/yyyy"));
        AssertFormattedCells(sheetName, ["A1"], ["3/15/2023"], [45000]);
    }

    [Fact]
    public void SetNumberFormat_ISODateFormat_DisplaysCorrectly()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1", [[45000]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", "yyyy-mm-dd"));
        AssertFormattedCells(sheetName, ["A1"], ["2023-03-15"], [45000]);
    }

    [Fact]
    public void SetNumberFormat_MultipleDates_AllDisplayCorrectly()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1:A3", [[45000], [45001], [45002]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "Z1", "mmm"));
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "Z1", [[45002]]));
        var month = Assert.Single(ReadFormattedCells(sheetName, ["Z1"])).Text;
        RequireSuccess(_rangeCommands.SetNumberFormats(batch, sheetName, "A1:A3",
            [["m/d/yyyy"], ["mm/dd/yyyy"], ["d-mmm-yyyy"]]));
        AssertFormattedCells(sheetName, ["A1", "A2", "A3"],
            ["3/15/2023", "03/16/2023", $"17-{month}-2023"], [45000, 45001, 45002]);
    }

    [Fact]
    public void SetNumberFormat_CurrencyFormat_NotAffected()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1", [[1234.56]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", "$#,##0.00"));
        AssertFormattedCells(sheetName, ["A1"], ["$1{group}234{decimal}56"], [1234.56]);
    }

    [Fact]
    public void SetNumberFormat_TimeFormat_DisplaysCorrectly()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1", [[0.75]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", "h:mm"));
        AssertFormattedCells(sheetName, ["A1"], ["18:00"], [0.75]);
    }

    [Fact]
    public void SetNumberFormat_DateTimeFormat_DisplaysCorrectly()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1", [[45000.75]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", "m/d/yyyy h:mm"));
        AssertFormattedCells(sheetName, ["A1"], ["3/15/2023 18:00"], [45000.75]);
    }

    [Theory]
    [InlineData("[NotAColor]0.00")]
    public void SetNumberFormat_InvalidCode_PreservesFormatsValuesAndDisplay(string format)
    {
        var batch = _fixture.BatchToken;
        var sheet = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheet, "A1:B1", [[12.5, 99.9]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheet, "A1:B1", "0.000"));
        var before = ReadFormattedCells(sheet, ["A1", "B1"]);
        var failure = Assert.Throws<InvalidOperationException>(() =>
            _rangeCommands.SetNumberFormat(batch, sheet, "A1", format));
        Assert.Contains("range.set-number-format failed", failure.Message);
        Assert.Equal(before, ReadFormattedCells(sheet, ["A1", "B1"]));
    }

    [Fact]
    public void SetNumberFormat_NonNumericCondition_MatchesNativeInterpretation()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(batch, sheetName, "A1:B1", [[12.5, 12.5]]));
        RequireSuccess(_rangeCommands.SetNumberFormat(batch, sheetName, "A1", "[>=not-number]0.00"));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            Microsoft.Office.Interop.Excel.Range? native = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[sheetName];
                native = sheet.Range["B1"];
                native.NumberFormatLocal = $"[>=not-number]0{context.FormatTranslator.DecimalSeparator}00";
            }
            finally
            {
                ComUtilities.Release(ref native);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var state = ReadFormattedCells(sheetName, ["A1", "B1"]);
        Assert.Equal(state[1], state[0]);
        Assert.Equal(12.5, state[0].Value);
    }

    private readonly record struct FormattedCell(string Format, string Text, double Value);

    private FormattedCell[] ReadFormattedCells(string sheetName, string[] addresses)
    {
        var reported = addresses.Select(address => Assert.Single(Assert.Single(
            RequireSuccess(_rangeCommands.GetNumberFormats(
                _fixture.BatchToken, sheetName, address)).Formats))).ToArray();
        var native = _fixture.ExecuteRawVerification((context, _) =>
        {
            Microsoft.Office.Interop.Excel.Sheets? sheets = null;
            Microsoft.Office.Interop.Excel.Worksheet? sheet = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Microsoft.Office.Interop.Excel.Worksheet)sheets[sheetName];
                return addresses.Select(address =>
                {
                    Microsoft.Office.Interop.Excel.Range? cell = null;
                    try
                    {
                        cell = sheet.Range[address];
                        cell.ColumnWidth = 50;
                        return new FormattedCell(Assert.IsType<string>(cell.NumberFormat),
                            Assert.IsType<string>(cell.Text), Assert.IsType<double>(cell.Value2));
                    }
                    finally { ComUtilities.Release(ref cell); }
                }).ToArray();
            }
            finally
            {
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        Assert.Equal(native.Select(cell => cell.Format), reported);
        return native;
    }

    private void AssertFormattedCells(string sheet, string[] addresses, string[] expected, double[] values)
    {
        var native = ReadFormattedCells(sheet, addresses);
        var separators = _fixture.ExecuteRawVerification((context, _) =>
            (context.FormatTranslator.DecimalSeparator, context.FormatTranslator.ThousandsSeparator));
        Assert.Equal(expected.Select(text => text
            .Replace("{decimal}", separators.DecimalSeparator, StringComparison.Ordinal)
            .Replace("{group}", separators.ThousandsSeparator, StringComparison.Ordinal)),
            native.Select(cell => cell.Text));
        Assert.Equal(values, native.Select(cell => cell.Value));
    }
}
