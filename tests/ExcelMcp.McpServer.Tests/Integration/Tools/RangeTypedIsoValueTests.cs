using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "true")]
public sealed class RangeTypedIsoValueTests : McpIntegrationTestBase
{
    private readonly string _testExcelFile;
    private string? _sessionId;

    public RangeTypedIsoValueTests(ITestOutputHelper output)
        : base(output, "RangeTypedIsoValueClient")
    {
        _testExcelFile = Path.Join(CreateTempDirectory("RangeTypedIsoValue"), "TypedIso.xlsx");
    }

    protected override async Task InitializeTestAsync()
    {
        _sessionId = await CreateWorkbookSessionAsync(_testExcelFile);
    }

    [Fact]
    public async Task SetValues_TypedIsoDate_RoundTripsThroughMcpProtocol()
    {
        var setValuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:B1",
            ["values"] = new object?[][]
            {
                [
                    new { type = "date", value = "2026-08-27", numberFormat = "@" },
                    "2026-08-27"
                ]
            }
        });
        AssertSetupSuccess(setValuesJson, "range.set-values typed ISO date");

        var getValuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:B1"
        });
        using var valuesDocument = JsonDocument.Parse(getValuesJson);
        Assert.True(valuesDocument.RootElement.GetProperty("success").GetBoolean(), getValuesJson);
        Assert.Equal(new DateTime(2026, 8, 27).ToOADate(), valuesDocument.RootElement.GetProperty("values")[0][0].GetDouble());
        Assert.Equal("2026-08-27", valuesDocument.RootElement.GetProperty("values")[0][1].GetString());

        var getFormatsJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-number-formats",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1:B1"
        });
        using var formatsDocument = JsonDocument.Parse(getFormatsJson);
        Assert.True(formatsDocument.RootElement.GetProperty("success").GetBoolean(), getFormatsJson);
        Assert.Equal("@", formatsDocument.RootElement.GetProperty("formats")[0][0].GetString(), StringComparer.OrdinalIgnoreCase);

        await CloseSessionAsync(_sessionId, save: false);
        _sessionId = null;
    }

    [Fact]
    public async Task SetValues_InvalidTypedIsoDate_ReturnsInvalidInputThroughMcpProtocol()
    {
        var setValuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["values"] = new object?[][]
            {
                [new { type = "datetime", value = "2026-08-27T10:30:00Z" }]
            }
        });

        using var document = ParseJsonResult(setValuesJson, "range.set-values invalid typed ISO date");
        AssertFailureEnvelope(
            document.RootElement,
            "range.set-values invalid typed ISO date",
            expectedExceptionType: "ArgumentException",
            expectedErrorCategory: "InvalidInput");
        Assert.Contains(
            "row 1, column 1",
            document.RootElement.GetProperty("errorMessage").GetString(),
            StringComparison.OrdinalIgnoreCase);

        await CloseSessionAsync(_sessionId, save: false);
        _sessionId = null;
    }
}
