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
public sealed class RangeFormulaErrorProtocolTests : McpIntegrationTestBase
{
    private static readonly string[][] ReferenceErrorFormula = [["=INDIRECT(\"A0\")"]];

    private readonly string _testExcelFile;
    private string? _sessionId;

    public RangeFormulaErrorProtocolTests(ITestOutputHelper output)
        : base(output, "RangeFormulaErrorClient")
    {
        _testExcelFile = Path.Join(
            CreateTempDirectory("RangeFormulaErrorProtocol"),
            "RangeFormulaError.xlsx");
    }

    protected override async Task InitializeTestAsync()
    {
        _sessionId = await CreateWorkbookSessionAsync(_testExcelFile);
    }

    [Fact]
    public async Task RangeReads_ReturnCanonicalFormulaErrorThroughMcp()
    {
        var setJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-formulas",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["formulas"] = ReferenceErrorFormula
        });
        AssertSetupSuccess(setJson, "range.set-formulas");

        var valuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1"
        });
        var formulasJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-formulas",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1"
        });

        AssertCanonicalReferenceError(valuesJson);
        AssertCanonicalReferenceError(formulasJson);

        await CloseSessionAsync(_sessionId, save: false);
        _sessionId = null;
    }

    private static void AssertCanonicalReferenceError(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean(), json);
        Assert.Equal("#REF!", root.GetProperty("values")[0][0].GetString());

        var error = root.GetProperty("cellErrors")[0];
        Assert.Equal("A1", error.GetProperty("cellAddress").GetString());
        Assert.Equal("#REF!", error.GetProperty("errorName").GetString());
        Assert.Equal("=INDIRECT(\"A0\")", error.GetProperty("formula").GetString());
        Assert.Equal(-2146826265, error.GetProperty("errorCode").GetInt32());
        Assert.Equal(-2146826265, error.GetProperty("currentValue").GetInt32());
    }
}
