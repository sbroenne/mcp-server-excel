using Sbroenne.ExcelMcp.ComInterop;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceRangeSpecializedTests
{
    private const int YellowFillColor = 65535;
    private const int CenterAlignment = -4108;
    private const int Issue585FillColor = 7949855;
    private const int WhiteFontColor = 16777215;

    [Fact]
    public void FormatRanges_AppliesSharedFormattingToEachTargetRange_AndLeavesUntargetedCellsUnchanged()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var untouchedBefore = new[]
        {
            ReadCellFormattingState(sheetName, "B1"),
            ReadCellFormattingState(sheetName, "B2")
        };

        var result = FormatRanges(
            batch,
            sheetName,
            ["A1:A2", "C1:C2"],
            bold: true,
            fillColor: "#FFFF00",
            horizontalAlignment: "center");

        Assert.True(result.Success, $"FormatRanges failed: {result.ErrorMessage}");
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(sheetName, "A1"));
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(sheetName, "A2"));
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(sheetName, "C1"));
        Assert.Equal(new CellFormattingState(true, YellowFillColor, CenterAlignment), ReadCellFormattingState(sheetName, "C2"));
        Assert.Equal(untouchedBefore[0], ReadCellFormattingState(sheetName, "B1"));
        Assert.Equal(untouchedBefore[1], ReadCellFormattingState(sheetName, "B2"));
    }

    [Fact]
    public void FormatRanges_WithNumberFormat_AppliesFormatCodeToAllTargetRanges()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var result = FormatRanges(
            batch,
            sheetName,
            ["A1:A2", "C1:C2"],
            numberFormat: "0.00%");

        Assert.True(result.Success, $"FormatRanges failed: {result.ErrorMessage}");
        Assert.Equal("0.00%", ReadCellNumberFormat(sheetName, "A1"));
        Assert.Equal("0.00%", ReadCellNumberFormat(sheetName, "A2"));
        Assert.Equal("0.00%", ReadCellNumberFormat(sheetName, "C1"));
        Assert.Equal("0.00%", ReadCellNumberFormat(sheetName, "C2"));
        Assert.NotEqual("0.00%", ReadCellNumberFormat(sheetName, "B1"));
    }

    [Fact]
    public void FormatRanges_InvalidTargetAddress_ErrorMessageIncludesIndex()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);

        var exception = Assert.Throws<ArgumentException>(() =>
            FormatRanges(
                batch,
                sheetName,
                ["A1:A2", "NotARange"],
                bold: true));

        Assert.Contains("1", exception.Message);
    }

    [Fact]
    public void FormatRanges_InvalidTargetAddress_FailsFast_AndDoesNotPartiallyApplyEarlierRanges()
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        var a1Before = ReadCellFormattingState(sheetName, "A1");
        var a2Before = ReadCellFormattingState(sheetName, "A2");

        var exception = Assert.Throws<ArgumentException>(() =>
            FormatRanges(
                batch,
                sheetName,
                ["A1:A2", "NotARange"],
                bold: true,
                fillColor: "#FFFF00",
                horizontalAlignment: "center"));

        Assert.Contains("range", exception.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(a1Before, ReadCellFormattingState(sheetName, "A1"));
        Assert.Equal(a2Before, ReadCellFormattingState(sheetName, "A2"));
    }

    [Fact]
    public async Task FormatRange_Issue585Payload_AppliesAndPersistsAfterReopen()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);

        var response = _fixture.Send(
            "rangeformat.format-range",
            new
            {
                sheetName,
                rangeAddress = "A1:J1",
                bold = true,
                fillColor = "#1F4E79",
                fontColor = "#FFFFFF"
            });

        Assert.True(response.Success, response.ErrorMessage);
        Assert.Equal(
            new Issue585FormattingState(true, Issue585FillColor, WhiteFontColor),
            ReadIssue585FormattingState(sheetName, "A1"));
        Assert.Equal(
            new Issue585FormattingState(true, Issue585FillColor, WhiteFontColor),
            ReadIssue585FormattingState(sheetName, "J1"));

        await _fixture.SaveAndReopenAsync();

        Assert.Equal(
            new Issue585FormattingState(true, Issue585FillColor, WhiteFontColor),
            ReadIssue585FormattingState(sheetName, "A1"));
        Assert.Equal(
            new Issue585FormattingState(true, Issue585FillColor, WhiteFontColor),
            ReadIssue585FormattingState(sheetName, "J1"));
    }

    [Fact]
    public async Task FormatRange_InvalidColor_ReturnsInvalidInputFailure()
    {
        var sheetName = _fixture.CreateTestSheet(_fixture.BatchToken);

        var response = await _fixture.SendForFailureAsync(
            "rangeformat.format-range",
            new
            {
                sheetName,
                rangeAddress = "A1:J1",
                fillColor = "not-a-color"
            });

        Assert.Equal("InvalidInput", response.ErrorCategory);
        Assert.Equal(nameof(ArgumentException), response.ExceptionType);
        Assert.Contains(
            "Invalid color format: not-a-color",
            response.ErrorMessage,
            StringComparison.Ordinal);
    }

    [Fact]
    public async Task FormatRange_MissingSheet_ReturnsComInteropFailure()
    {
        var response = await _fixture.SendForFailureAsync(
            "rangeformat.format-range",
            new
            {
                sheetName = $"Missing_{Guid.NewGuid():N}",
                rangeAddress = "A1:J1",
                bold = true,
                fillColor = "#1F4E79",
                fontColor = "#FFFFFF"
            });

        Assert.Equal("ComInterop", response.ErrorCategory);
        Assert.Equal("COMException", response.ExceptionType);
        Assert.Contains(
            "Invalid index",
            response.ErrorMessage,
            StringComparison.OrdinalIgnoreCase);
    }

    private Sbroenne.ExcelMcp.Core.Models.OperationResult FormatRanges(
        Sbroenne.ExcelMcp.ComInterop.Session.IExcelBatch batch,
        string sheetName,
        string[] rangeAddresses,
        string? fontName = null,
        double? fontSize = null,
        bool? bold = null,
        bool? italic = null,
        bool? underline = null,
        string? fontColor = null,
        string? fillColor = null,
        string? borderStyle = null,
        string? borderColor = null,
        string? borderWeight = null,
        string? horizontalAlignment = null,
        string? verticalAlignment = null,
        bool? wrapText = null,
        int? orientation = null,
        string? numberFormat = null) =>
        _commands.FormatRanges(
            batch,
            sheetName,
            rangeAddresses,
            fontName,
            fontSize,
            bold,
            italic,
            underline,
            fontColor,
            fillColor,
            borderStyle,
            borderColor,
            borderWeight,
            horizontalAlignment,
            verticalAlignment,
            wrapText,
            orientation,
            numberFormat);

    private CellFormattingState ReadCellFormattingState(
        string sheetName,
        string cellAddress) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? font = null;
            dynamic? interior = null;
            try
            {
                sheet = context.Book.Worksheets[sheetName];
                range = sheet.Range[cellAddress];
                font = range.Font;
                interior = range.Interior;
                return new CellFormattingState(
                    Convert.ToBoolean(font.Bold),
                    Convert.ToInt32(interior.Color),
                    Convert.ToInt32(range.HorizontalAlignment));
            }
            finally
            {
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });

    private string ReadCellNumberFormat(string sheetName, string cellAddress) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                sheet = context.Book.Worksheets[sheetName];
                range = sheet.Range[cellAddress];
                return (string)(((Microsoft.Office.Interop.Excel.Range)range).NumberFormat ?? "General");
            }
            finally
            {
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });

    private Issue585FormattingState ReadIssue585FormattingState(
        string sheetName,
        string cellAddress) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? font = null;
            dynamic? interior = null;
            try
            {
                sheet = context.Book.Worksheets[sheetName];
                range = sheet.Range[cellAddress];
                font = range.Font;
                interior = range.Interior;
                return new Issue585FormattingState(
                    Convert.ToBoolean(font.Bold),
                    Convert.ToInt32(interior.Color),
                    Convert.ToInt32(font.Color));
            }
            finally
            {
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });

    private readonly record struct CellFormattingState(
        bool Bold,
        int FillColor,
        int HorizontalAlignment);

    private readonly record struct Issue585FormattingState(
        bool Bold,
        int FillColor,
        int FontColor);
}
