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
        Assert.True(_rangeCommands.SetValues(batch, sheetName, "A1:B1", [[45000.75, 45000.75]]).Success);
        var written = _rangeCommands.SetNumberFormat(batch, sheetName, "A1", format);
        Assert.True(written.Success, written.ErrorMessage);
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
                Assert.Equal(native.Text, actual.Text);
                Assert.Equal(native.NumberFormat, actual.NumberFormat);
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
        Assert.True(_rangeCommands.SetValues(batch, sheetName, "A1:B1", [[1234.56, 1234.56]]).Success);
        var written = _rangeCommands.SetNumberFormat(batch, sheetName, "A1", "$#,##0.00");
        Assert.True(written.Success, written.ErrorMessage);
        var read = _rangeCommands.GetNumberFormats(batch, sheetName, "A1");
        Assert.True(read.Success, read.ErrorMessage);
        var returnedFormat = Assert.Single(Assert.Single(read.Formats));
        var rewritten = _rangeCommands.SetNumberFormat(batch, sheetName, "A1", returnedFormat);
        Assert.True(rewritten.Success, rewritten.ErrorMessage);
        var reread = _rangeCommands.GetNumberFormats(batch, sheetName, "A1");
        Assert.True(reread.Success, reread.ErrorMessage);
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
        Assert.True(_rangeCommands.SetValues(batch, sheetName, address, data).Success);
        if (action is "column" or "grid")
        {
            var written = _rangeCommands.SetNumberFormats(batch, sheetName, address, codes);
            Assert.True(written.Success, written.ErrorMessage);
        }
        else
        {
            for (var index = 0; index < formats.Length; index++)
            {
                var cellAddress = $"A{index + 1}";
                var written = action == "single"
                    ? _rangeCommands.SetNumberFormat(batch, sheetName, cellAddress, formats[index])
                    : _rangeCommands.FormatRanges(batch, sheetName, [cellAddress],
                        null, null, null, null, null, null, null, null, null, null, null, null, null, null, formats[index]);
                Assert.True(written.Success, written.ErrorMessage);
            }
        }
        var read = _rangeCommands.GetNumberFormats(batch, sheetName, address);
        Assert.True(read.Success, read.ErrorMessage);
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
            _rangeCommands.SetValues(batch, sheetName, "A1", [[0.125]]);
            _rangeCommands.SetNumberFormat(batch, sheetName, "A1", "0.00%");
            var result = _rangeCommands.GetNumberFormats(batch, sheetName, "A1");
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
        _rangeCommands.SetValues(batch, sheetName, "A1:A2", [[1.25], [1.75]]);
        _rangeCommands.SetNumberFormat(batch, sheetName, "A1:A2", "[>=1.5]\"high\";\"low\"");

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
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Log translator info
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            _output.WriteLine($"FormatTranslator: {ctx.FormatTranslator}");
        });

        // Set a date value (45000 = March 15, 2023)
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];
            cell.Value2 = 45000; // March 15, 2023
        });

        // Act - Set format using US format code "m/d/yyyy"
        var result = _rangeCommands.SetNumberFormat(batch, sheetName, "A1", "m/d/yyyy");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        // Verify the display is correct (not "0/d/yyyy" or other broken formats)
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];

            string displayedText = cell.Text?.ToString() ?? "null";
            string appliedFormat = cell.NumberFormat?.ToString() ?? "null";
            string localFormat = cell.NumberFormatLocal?.ToString() ?? "null";

            _output.WriteLine($"Set US format 'm/d/yyyy':");
            _output.WriteLine($"  Applied NumberFormat: '{appliedFormat}'");
            _output.WriteLine($"  NumberFormatLocal: '{localFormat}'");
            _output.WriteLine($"  Displayed text: '{displayedText}'");

            // The display should contain the date parts (3, 15, 2023 or 15, 3, 2023)
            // NOT "0/d/yyyy" which happens when 'm' is misinterpreted as minutes (=0)
            Assert.DoesNotContain("0/", displayedText);
            Assert.DoesNotContain("/0/", displayedText);

            // Should contain year 2023 (45000 = March 15, 2023)
            Assert.Contains("2023", displayedText);
        });
    }

    [Fact]
    public void SetNumberFormat_ISODateFormat_DisplaysCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set a date value (45000 = March 15, 2023)
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];
            cell.Value2 = 45000; // March 15, 2023
        });

        // Act - Set format using ISO format "yyyy-mm-dd"
        var result = _rangeCommands.SetNumberFormat(batch, sheetName, "A1", "yyyy-mm-dd");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            dynamic cell = sheet.Range["A1"];

            string displayedText = cell.Text?.ToString() ?? "null";

            _output.WriteLine($"Set ISO format 'yyyy-mm-dd':");
            _output.WriteLine($"  Displayed text: '{displayedText}'");

            // Should display as ISO format: 2023-03-15
            Assert.Equal("2023-03-15", displayedText);
        });
    }

    [Fact]
    public void SetNumberFormat_MultipleDates_AllDisplayCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set date values in A1:A3
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            sheet.Range["A1"].Value2 = 45000; // March 15, 2023
            sheet.Range["A2"].Value2 = 45001; // March 16, 2023
            sheet.Range["A3"].Value2 = 45002; // March 17, 2023
        });

        // Act - Set all three cells with different date formats
        var formats = new List<List<string>>
        {
            new() { "m/d/yyyy" },
            new() { "mm/dd/yyyy" },
            new() { "d-mmm-yyyy" }
        };

        var result = _rangeCommands.SetNumberFormats(batch, sheetName, "A1:A3", formats);

        // Assert
        Assert.True(result.Success, $"SetNumberFormats failed: {result.ErrorMessage}");

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];

            var texts = new[]
            {
                sheet.Range["A1"].Text?.ToString() ?? "null",
                sheet.Range["A2"].Text?.ToString() ?? "null",
                sheet.Range["A3"].Text?.ToString() ?? "null"
            };

            _output.WriteLine($"A1 (m/d/yyyy): '{texts[0]}'");
            _output.WriteLine($"A2 (mm/dd/yyyy): '{texts[1]}'");
            _output.WriteLine($"A3 (d-mmm-yyyy): '{texts[2]}'");

            // All should contain the year 2023, not broken format codes
            foreach (var text in texts)
            {
                Assert.Contains("2023", text);
                Assert.DoesNotContain("0/d", text);
            }
        });
    }

    [Fact]
    public void SetNumberFormat_CurrencyFormat_NotAffected()
    {
        // Currency formats should NOT be affected by date translation

        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set a currency value
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            sheet.Range["A1"].Value2 = 1234.56;
        });

        // Act - Set currency format
        var result = _rangeCommands.SetNumberFormat(batch, sheetName, "A1", "$#,##0.00");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            string displayedText = sheet.Range["A1"].Text?.ToString() ?? "null";

            _output.WriteLine($"Currency format '$#,##0.00': '{displayedText}'");

            // Should contain the dollar sign and proper formatting
            Assert.Equal(
                $"$1{ctx.FormatTranslator.ThousandsSeparator}234{ctx.FormatTranslator.DecimalSeparator}56",
                displayedText);
            Assert.Equal(1234.56, Convert.ToDouble(sheet.Range["A1"].Value2));
        });
    }

    [Fact]
    public void SetNumberFormat_TimeFormat_DisplaysCorrectly()
    {
        // Arrange
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set a time value (0.75 = 6:00 PM / 18:00)
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            sheet.Range["A1"].Value2 = 0.75; // 18:00
        });

        // Act - Set time format
        var result = _rangeCommands.SetNumberFormat(batch, sheetName, "A1", "h:mm");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            string displayedText = sheet.Range["A1"].Text?.ToString() ?? "null";

            _output.WriteLine($"Time format 'h:mm': '{displayedText}'");

            // Should show time (18:00 or 6:00 PM depending on locale)
            Assert.True(displayedText.Contains("18:00") || displayedText.Contains("6:00"),
                $"Expected time display, got '{displayedText}'");
        });
    }

    [Fact]
    public void SetNumberFormat_DateTimeFormat_DisplaysCorrectly()
    {
        // Test combined date+time format

        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        // Set a date+time value (45000.75 = March 15, 2023 at 18:00)
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            sheet.Range["A1"].Value2 = 45000.75;
        });

        // Act - Set date+time format
        var result = _rangeCommands.SetNumberFormat(batch, sheetName, "A1", "m/d/yyyy h:mm");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[sheetName];
            string displayedText = sheet.Range["A1"].Text?.ToString() ?? "null";

            _output.WriteLine($"DateTime format 'm/d/yyyy h:mm': '{displayedText}'");

            // Should contain year (date part works)
            Assert.Contains("2023", displayedText);

            // Should contain time separator (time part works)
            Assert.Contains(":", displayedText);
        });
    }
}
