using Sbroenne.ExcelMcp.ComInterop.Formatting;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

/// <summary>
/// Tests that date format translation works correctly across locales.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "Core")]
[Trait("Feature", "Ranges")]
[Trait("RequiresExcel", "true")]
public class DateFormatTranslationTests : IClassFixture<RangeTestsFixture>
{
    private readonly ITestOutputHelper _output;
    private readonly RangeTestsFixture _fixture;
    private readonly RangeCommands _rangeCommands;

    public DateFormatTranslationTests(RangeTestsFixture fixture, ITestOutputHelper output)
    {
        _fixture = fixture;
        _output = output;
        _rangeCommands = new RangeCommands();
    }

    [Fact]
    public void SetNumberFormat_USDateFormat_DisplaysCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Log translator info
        batch.Execute((ctx, ct) =>
        {
            _output.WriteLine($"DateFormatter: {ctx.DateFormatter}");
        });

        // Set a date value (45000 = March 15, 2023)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            dynamic cell = sheet.Range["A1"];
            cell.Value2 = 45000; // March 15, 2023
        });

        // Act - Set format using US format code "m/d/yyyy"
        var result = _rangeCommands.SetNumberFormat(batch, "Sheet1", "A1", "m/d/yyyy");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        // Verify the display is correct (not "0/d/yyyy" or other broken formats)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
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
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set a date value (45000 = March 15, 2023)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            dynamic cell = sheet.Range["A1"];
            cell.Value2 = 45000; // March 15, 2023
        });

        // Act - Set format using ISO format "yyyy-mm-dd"
        var result = _rangeCommands.SetNumberFormat(batch, "Sheet1", "A1", "yyyy-mm-dd");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
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
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set date values in A1:A3
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
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

        var result = _rangeCommands.SetNumberFormats(batch, "Sheet1", "A1:A3", formats);

        // Assert
        Assert.True(result.Success, $"SetNumberFormats failed: {result.ErrorMessage}");

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);

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

        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set a currency value
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = 1234.56;
        });

        // Act - Set currency format
        var result = _rangeCommands.SetNumberFormat(batch, "Sheet1", "A1", "$#,##0.00");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            string displayedText = sheet.Range["A1"].Text?.ToString() ?? "null";

            _output.WriteLine($"Currency format '$#,##0.00': '{displayedText}'");

            // Should contain the dollar sign and proper formatting
            Assert.Contains("$", displayedText);
            Assert.Contains("1,234.56", displayedText);
        });
    }

    [Fact]
    public void SetNumberFormat_TimeFormat_DisplaysCorrectly()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set a time value (0.75 = 6:00 PM / 18:00)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = 0.75; // 18:00
        });

        // Act - Set time format
        var result = _rangeCommands.SetNumberFormat(batch, "Sheet1", "A1", "h:mm");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
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

        var testFile = _fixture.CreateTestFile();

        using var batch = ExcelSession.BeginBatch(testFile);

        // Set a date+time value (45000.75 = March 15, 2023 at 18:00)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            sheet.Range["A1"].Value2 = 45000.75;
        });

        // Act - Set date+time format
        var result = _rangeCommands.SetNumberFormat(batch, "Sheet1", "A1", "m/d/yyyy h:mm");

        // Assert
        Assert.True(result.Success, $"SetNumberFormat failed: {result.ErrorMessage}");

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Item(1);
            string displayedText = sheet.Range["A1"].Text?.ToString() ?? "null";

            _output.WriteLine($"DateTime format 'm/d/yyyy h:mm': '{displayedText}'");

            // Should contain year (date part works)
            Assert.Contains("2023", displayedText);

            // Should contain time separator (time part works)
            Assert.Contains(":", displayedText);
        });
    }
}
