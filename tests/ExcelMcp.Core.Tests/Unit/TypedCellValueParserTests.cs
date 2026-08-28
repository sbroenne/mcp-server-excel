using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "Range")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class TypedCellValueParserTests
{
    [Theory]
    [InlineData("""{"type":"date","value":"2026-08-27"}""", "2026-08-27T00:00:00.0000000", "yyyy-mm-dd")]
    [InlineData("""{"type":"datetime","value":"2024-02-29T12:34:56.25"}""", "2024-02-29T12:34:56.2500000", "yyyy-mm-dd hh:mm:ss")]
    [InlineData("""{"type":"datetime-offset","value":"2026-08-27T03:15:00-04:00"}""", "2026-08-27T07:15:00.0000000", "yyyy-mm-dd hh:mm:ss")]
    [InlineData("""{"type":"datetime-offset","value":"2026-08-27T07:15:00Z"}""", "2026-08-27T07:15:00.0000000", "yyyy-mm-dd hh:mm:ss")]
    public void Parse_ValidTypedValue_ReturnsDeterministicDateTime(
        string json,
        string expectedDateTime,
        string expectedNumberFormat)
    {
        using var document = JsonDocument.Parse(json);

        var result = TypedCellValueParser.Parse(document.RootElement, 1, 1);

        Assert.True(result.IsTypedDate);
        Assert.Equal(DateTime.Parse(expectedDateTime, System.Globalization.CultureInfo.InvariantCulture), result.DateTimeValue);
        Assert.Equal(DateTimeKind.Unspecified, result.DateTimeValue.Kind);
        Assert.Equal(expectedNumberFormat, result.NumberFormat);
    }

    [Fact]
    public void Parse_CustomNumberFormat_PreservesCallerFormat()
    {
        using var document = JsonDocument.Parse(
            """{"type":"date","value":"2026-08-27","numberFormat":"dd-mmm-yyyy"}""");

        var result = TypedCellValueParser.Parse(document.RootElement, 2, 3);

        Assert.Equal("dd-mmm-yyyy", result.NumberFormat);
    }

    [Fact]
    public void Parse_CoreTypedModel_UsesSameContractAsJson()
    {
        var input = new TypedCellValue
        {
            Type = TypedCellValueType.DateTime,
            Value = "2026-08-27T14:30:00"
        };

        var result = TypedCellValueParser.Parse(input, 1, 1);

        Assert.True(result.IsTypedDate);
        Assert.Equal(new DateTime(2026, 8, 27, 14, 30, 0), result.DateTimeValue);
        Assert.Equal("yyyy-mm-dd hh:mm:ss", result.NumberFormat);
    }

    [Fact]
    public void Parse_OrdinaryIsoString_RemainsAnOrdinaryString()
    {
        const string value = "2026-08-27";

        var result = TypedCellValueParser.Parse(value, 1, 1);

        Assert.False(result.IsTypedDate);
        Assert.Equal(value, result.PrimitiveValue);
    }

    [Theory]
    [InlineData("""{"type":"date","value":"2026-08-27T10:30:00"}""", "date", "yyyy-MM-dd")]
    [InlineData("""{"type":"datetime","value":"2026-08-27T10:30:00Z"}""", "datetime", "without a timezone")]
    [InlineData("""{"type":"datetime-offset","value":"2026-08-27T10:30:00"}""", "datetime-offset", "with Z or an offset")]
    [InlineData("""{"type":"date","value":"2026-02-29"}""", "date", "yyyy-MM-dd")]
    [InlineData("""{"type":"timestamp","value":"2026-08-27"}""", "timestamp", "date, datetime, datetime-offset")]
    [InlineData("""{"type":"date"}""", "value", "required")]
    [InlineData("""{"type":"date","value":"2026-08-27","numberFormat":" "}""", "numberFormat", "must not be empty")]
    public void Parse_InvalidTypedValue_ThrowsCoordinateAwareError(
        string json,
        string expectedMessagePart1,
        string expectedMessagePart2)
    {
        using var document = JsonDocument.Parse(json);

        var exception = Assert.Throws<ArgumentException>(
            () => TypedCellValueParser.Parse(document.RootElement, 4, 2));

        Assert.Contains("row 4, column 2", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(expectedMessagePart1, exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(expectedMessagePart2, exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("1900-01-01T00:00:00", false, 1.0)]
    [InlineData("1900-02-28T00:00:00", false, 59.0)]
    [InlineData("1900-03-01T00:00:00", false, 61.0)]
    [InlineData("1904-01-01T00:00:00", true, 0.0)]
    [InlineData("2024-02-29T12:00:00", false, 45351.5)]
    [InlineData("2024-02-29T12:00:00", true, 43889.5)]
    public void ToExcelSerial_UsesWorkbookDateSystem(
        string dateTime,
        bool use1904DateSystem,
        double expected)
    {
        var value = DateTime.Parse(dateTime, System.Globalization.CultureInfo.InvariantCulture);

        var result = TypedCellValueParser.ToExcelSerial(value, use1904DateSystem, 1, 1);

        Assert.Equal(expected, result);
    }

    [Fact]
    public void ToExcelSerial_DateBeforeWorkbookEpoch_ThrowsCoordinateAwareError()
    {
        var exception = Assert.Throws<ArgumentException>(
            () => TypedCellValueParser.ToExcelSerial(new DateTime(1903, 12, 31), true, 3, 5));

        Assert.Contains("row 3, column 5", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("1904 date system", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("1904-01-01", exception.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("2026-08-27T12:34:56.1234567")]
    [InlineData("2026-08-27T06:32:27.0483540")]
    public void ToExcelSerial_PreservesSevenDigitFractionalSeconds(string input)
    {
        var value = DateTime.ParseExact(
            input,
            "yyyy-MM-dd'T'HH:mm:ss.FFFFFFF",
            System.Globalization.CultureInfo.InvariantCulture);

        var serial = TypedCellValueParser.ToExcelSerial(value, false, 1, 1);
        var reconstructedTimeTicks = (long)Math.Round(
            (serial - Math.Truncate(serial)) * TimeSpan.TicksPerDay,
            MidpointRounding.AwayFromZero);

        Assert.InRange(
            Math.Abs(reconstructedTimeTicks - value.TimeOfDay.Ticks),
            0,
            5);

        var oleDateTicks = (long)Math.Round(
            (value.ToOADate() - Math.Truncate(value.ToOADate())) * TimeSpan.TicksPerDay,
            MidpointRounding.AwayFromZero);
        Assert.True(
            Math.Abs(oleDateTicks - value.TimeOfDay.Ticks) > 1_000,
            "The test value must demonstrate the sub-millisecond loss in DateTime.ToOADate().");
    }
}
