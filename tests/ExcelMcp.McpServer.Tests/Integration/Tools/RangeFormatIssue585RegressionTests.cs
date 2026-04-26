using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class RangeFormatIssue585RegressionTests : McpIntegrationTestBase
{
    private const string IssueSheetName = "Toutes les transactions";
    private const int IssueFillColor = 7949855;
    private const int WhiteFontColor = 16777215;

    private readonly string _testExcelFile;
    private string? _sessionId;

    public RangeFormatIssue585RegressionTests(ITestOutputHelper output)
        : base(output, "RangeFormatIssue585RegressionClient")
    {
        _testExcelFile = Path.Join(CreateTempDirectory("RangeFormatIssue585"), "RangeFormatIssue585.xlsx");
    }

    protected override async Task InitializeTestAsync()
    {
        _sessionId = await CreateWorkbookSessionAsync(_testExcelFile);
    }

    [Fact]
    public async Task FormatRange_MinimalBoldPayload_AppliesViaMcpProtocol()
    {
        await CreateIssueSheetAsync();

        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-range",
            ["session_id"] = _sessionId,
            ["sheet_name"] = IssueSheetName,
            ["range_address"] = "A1:J1",
            ["bold"] = true
        });
        AssertSuccess(formatJson, "range_format.format-range minimal");

        await CloseTrackedSessionAsync(save: true);

        using var batch = ExcelSession.BeginBatch(_testExcelFile);
        Assert.True(ReadCellFormatting(batch, IssueSheetName, "A1").Bold);
        Assert.True(ReadCellFormatting(batch, IssueSheetName, "J1").Bold);
    }

    [Fact]
    public async Task FormatRange_Issue585Payload_AppliesViaMcpProtocol()
    {
        await CreateIssueSheetAsync();

        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-range",
            ["session_id"] = _sessionId,
            ["sheet_name"] = IssueSheetName,
            ["range_address"] = "A1:J1",
            ["bold"] = true,
            ["fill_color"] = "#1F4E79",
            ["font_color"] = "#FFFFFF"
        });
        AssertSuccess(formatJson, "range_format.format-range issue-585");

        await CloseTrackedSessionAsync(save: true);

        using var batch = ExcelSession.BeginBatch(_testExcelFile);
        var a1 = ReadCellFormatting(batch, IssueSheetName, "A1");
        var j1 = ReadCellFormatting(batch, IssueSheetName, "J1");

        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), a1);
        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), j1);
    }

    [Fact]
    public async Task FormatRange_Issue585Payload_AppliesViaMcpProtocol_AfterReopeningWorkbook()
    {
        await CreateIssueSheetAsync();
        await CloseTrackedSessionAsync(save: true);

        _sessionId = await OpenWorkbookSessionAsync(_testExcelFile);

        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-range",
            ["session_id"] = _sessionId,
            ["sheet_name"] = IssueSheetName,
            ["range_address"] = "A1:J1",
            ["bold"] = true,
            ["fill_color"] = "#1F4E79",
            ["font_color"] = "#FFFFFF"
        });
        AssertSuccess(formatJson, "range_format.format-range issue-585 reopen");

        await CloseTrackedSessionAsync(save: true);

        using var batch = ExcelSession.BeginBatch(_testExcelFile);
        var a1 = ReadCellFormatting(batch, IssueSheetName, "A1");
        var j1 = ReadCellFormatting(batch, IssueSheetName, "J1");

        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), a1);
        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), j1);
    }

    [Fact]
    public async Task FormatRange_Issue585Payload_WithExplicitNullOptionals_AppliesViaMcpProtocol()
    {
        await CreateIssueSheetAsync();

        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-range",
            ["session_id"] = _sessionId,
            ["sheet_name"] = IssueSheetName,
            ["range_address"] = "A1:J1",
            ["font_name"] = null,
            ["font_size"] = null,
            ["bold"] = true,
            ["italic"] = null,
            ["underline"] = null,
            ["font_color"] = "#FFFFFF",
            ["fill_color"] = "#1F4E79",
            ["border_style"] = null,
            ["border_color"] = null,
            ["border_weight"] = null,
            ["horizontal_alignment"] = null,
            ["vertical_alignment"] = null,
            ["wrap_text"] = null,
            ["orientation"] = null
        });
        AssertSuccess(formatJson, "range_format.format-range issue-585 explicit-nulls");

        await CloseTrackedSessionAsync(save: true);

        using var batch = ExcelSession.BeginBatch(_testExcelFile);
        var a1 = ReadCellFormatting(batch, IssueSheetName, "A1");
        var j1 = ReadCellFormatting(batch, IssueSheetName, "J1");

        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), a1);
        Assert.Equal(new CellFormattingState(true, IssueFillColor, WhiteFontColor), j1);
    }

    [Fact]
    public async Task FormatRange_InvalidColor_ReturnsTransparentFailureEnvelopeViaMcpProtocol()
    {
        await CreateIssueSheetAsync();

        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-range",
            ["session_id"] = _sessionId,
            ["sheet_name"] = IssueSheetName,
            ["range_address"] = "A1:J1",
            ["fill_color"] = "not-a-color"
        });

        using var formatDoc = ParseJsonResult(formatJson, "range_format.format-range invalid-color");
        AssertFailureEnvelope(
            formatDoc.RootElement,
            "range_format.format-range invalid-color",
            expectedExceptionType: "ArgumentException",
            expectedErrorCategory: "InvalidInput");

        Assert.Equal("rangeformat.format-range", formatDoc.RootElement.GetProperty("command").GetString());
        Assert.Equal(_sessionId, formatDoc.RootElement.GetProperty("sessionId").GetString());

        var errorMessage = formatDoc.RootElement.GetProperty("errorMessage").GetString();
        Assert.Contains("Invalid color format: not-a-color", errorMessage, StringComparison.Ordinal);
    }

    [Fact]
    public async Task FormatRange_Issue585Payload_OnMissingSheet_ReturnsTransparentFailure()
    {
        var formatJson = await CallToolAsync("range_format", new Dictionary<string, object?>
        {
            ["action"] = "format-range",
            ["session_id"] = _sessionId,
            ["sheet_name"] = IssueSheetName,
            ["range_address"] = "A1:J1",
            ["bold"] = true,
            ["fill_color"] = "#1F4E79",
            ["font_color"] = "#FFFFFF"
        });

        using var doc = ParseJsonResult(formatJson, "range_format.format-range missing-sheet");
        AssertFailureEnvelope(
            doc.RootElement,
            "range_format.format-range missing-sheet",
            expectedExceptionType: "COMException",
            expectedErrorCategory: "ComInterop",
            expectedHResult: string.Empty);

        Assert.Equal("rangeformat.format-range", doc.RootElement.GetProperty("command").GetString());
        Assert.Equal(_sessionId, doc.RootElement.GetProperty("sessionId").GetString());

        AssertHasNonEmptyStringProperty(doc.RootElement, "hresult", "range_format.format-range missing-sheet");

        var errorMessage = doc.RootElement.GetProperty("errorMessage").GetString();
        Assert.NotNull(errorMessage);
        Assert.Contains("COMException", errorMessage, StringComparison.Ordinal);
        Assert.Contains("Invalid index", errorMessage, StringComparison.OrdinalIgnoreCase);
    }

    private async Task CreateIssueSheetAsync()
    {
        await CreateWorksheetAsync(_sessionId!, IssueSheetName);
    }

    private async Task<string> OpenWorkbookSessionAsync(string workbookPath)
    {
        var openJson = await CallToolAsync("file", new Dictionary<string, object?>
        {
            ["action"] = "open",
            ["path"] = workbookPath
        });

        AssertSetupSuccess(openJson, $"file.open ({Path.GetFileName(workbookPath)})");

        using var openDoc = ParseJsonResult(openJson, $"file.open ({Path.GetFileName(workbookPath)})");
        var sessionId = openDoc.RootElement.GetProperty("session_id").GetString();
        TrackSession(sessionId);
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        return sessionId!;
    }

    private async Task CloseTrackedSessionAsync(bool save)
    {
        await CloseSessionAsync(_sessionId, save);
        _sessionId = null;
    }

    private static CellFormattingState ReadCellFormatting(IExcelBatch batch, string sheetName, string cellAddress)
    {
        return batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? range = null;
            dynamic? font = null;
            dynamic? interior = null;

            try
            {
                sheet = ctx.Book.Worksheets[sheetName];
                range = sheet.Range[cellAddress];
                font = range.Font;
                interior = range.Interior;

                return new CellFormattingState(
                    Bold: Convert.ToBoolean(font.Bold),
                    FillColor: interior.Color == null ? null : Convert.ToInt32(interior.Color),
                    FontColor: font.Color == null ? null : Convert.ToInt32(font.Color));
            }
            finally
            {
                ComUtilities.Release(ref interior!);
                ComUtilities.Release(ref font!);
                ComUtilities.Release(ref range!);
                ComUtilities.Release(ref sheet!);
            }
        });
    }

    private sealed record CellFormattingState(bool Bold, int? FillColor, int? FontColor);
}
