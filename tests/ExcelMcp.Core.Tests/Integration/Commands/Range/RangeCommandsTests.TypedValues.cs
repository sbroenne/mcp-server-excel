using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Range;

public partial class RangeCommandsTests
{
    [Fact]
    public void SetValues_TypedIsoValues_WriteSerialsFormatsAndPreservePlainStrings()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        using var document = JsonDocument.Parse(
            """
            [[
              {"type":"date","value":"2026-08-27"},
              {"type":"datetime","value":"2024-02-29T12:34:56.25"},
              {"type":"datetime-offset","value":"2026-08-27T03:15:00-04:00","numberFormat":"@"},
              "2026-08-27",
              "00123"
            ]]
            """);
        var values = document.RootElement[0]
            .EnumerateArray()
            .Select(cell => (object?)cell.Clone())
            .ToList();
        _commands.SetNumberFormat(batch, sheetName, "D1:E1", "0.00");
        var expectedDateFormat = batch.Execute((ctx, ct) =>
            ctx.FormatTranslator.TranslateToLocale("yyyy-mm-dd"));
        var expectedDateTimeFormat = batch.Execute((ctx, ct) =>
            ctx.FormatTranslator.TranslateToLocale("yyyy-mm-dd hh:mm:ss"));
        var expectedPrimitiveFormat = batch.Execute((ctx, ct) =>
            ctx.FormatTranslator.TranslateToLocale("0.00"));

        var writeResult = _commands.SetValues(batch, sheetName, "A1:E1", [values]);

        Assert.True(writeResult.Success, writeResult.ErrorMessage);
        var readResult = _commands.GetValues(batch, sheetName, "A1:E1");
        Assert.True(readResult.Success, readResult.ErrorMessage);
        Assert.Equal(new DateTime(2026, 8, 27).ToOADate(), Convert.ToDouble(readResult.Values[0][0], CultureInfo.InvariantCulture));
        Assert.Equal(new DateTime(2024, 2, 29, 12, 34, 56, 250).ToOADate(), Convert.ToDouble(readResult.Values[0][1], CultureInfo.InvariantCulture), 10);
        Assert.Equal(new DateTime(2026, 8, 27, 7, 15, 0).ToOADate(), Convert.ToDouble(readResult.Values[0][2], CultureInfo.InvariantCulture), 10);
        Assert.Equal("2026-08-27", readResult.Values[0][3]);
        Assert.Equal("00123", readResult.Values[0][4]);

        var formatResult = _commands.GetNumberFormats(batch, sheetName, "A1:E1");
        Assert.True(formatResult.Success, formatResult.ErrorMessage);
        Assert.Equal(expectedDateFormat, formatResult.Formats[0][0], StringComparer.OrdinalIgnoreCase);
        Assert.Equal(expectedDateTimeFormat, formatResult.Formats[0][1], StringComparer.OrdinalIgnoreCase);
        Assert.Equal("@", formatResult.Formats[0][2], StringComparer.OrdinalIgnoreCase);
        Assert.Equal(expectedPrimitiveFormat, formatResult.Formats[0][3], StringComparer.OrdinalIgnoreCase);
        Assert.Equal(expectedPrimitiveFormat, formatResult.Formats[0][4], StringComparer.OrdinalIgnoreCase);
    }

    [Fact]
    public void SetValues_TypedDate_Uses1904WorkbookDateSystem()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        try
        {
            batch.Execute((ctx, ct) =>
            {
                ctx.Book.Date1904 = true;
                return true;
            });
            using var document = JsonDocument.Parse("""[[{"type":"date","value":"1904-01-01"}]]""");
            var values = new List<List<object?>> { new() { document.RootElement[0][0].Clone() } };

            var writeResult = _commands.SetValues(batch, sheetName, "A1", values);

            Assert.True(writeResult.Success, writeResult.ErrorMessage);
            var readResult = _commands.GetValues(batch, sheetName, "A1");
            Assert.True(readResult.Success, readResult.ErrorMessage);
            Assert.Equal(0.0, Convert.ToDouble(readResult.Values[0][0], CultureInfo.InvariantCulture));
        }
        finally
        {
            batch.Execute((ctx, ct) =>
            {
                ctx.Book.Date1904 = false;
                return true;
            });
        }
    }

    [Fact]
    public void SetValues_InvalidTypedValue_FailsBeforeExcelWrite()
    {
        using var batch = ExcelSession.BeginBatch(_fixture.TestFilePath);
        var sheetName = _fixture.CreateTestSheet(batch);
        _commands.SetValues(batch, sheetName, "A1", [["unchanged"]]);
        using var document = JsonDocument.Parse("""[[{"type":"datetime","value":"2026-08-27T10:30:00Z"}]]""");
        var values = new List<List<object?>> { new() { document.RootElement[0][0].Clone() } };

        var exception = Assert.Throws<ArgumentException>(
            () => _commands.SetValues(batch, sheetName, "A1", values));

        Assert.Contains("row 1, column 1", exception.Message, StringComparison.OrdinalIgnoreCase);
        var readResult = _commands.GetValues(batch, sheetName, "A1");
        Assert.Equal("unchanged", readResult.Values[0][0]);
    }
}
