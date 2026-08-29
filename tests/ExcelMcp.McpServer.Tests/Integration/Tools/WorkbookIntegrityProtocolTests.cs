using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Workbook")]
[Trait("RequiresExcel", "true")]
public sealed class WorkbookIntegrityProtocolTests : McpIntegrationTestBase
{
    private static readonly string[][] ReferenceErrorFormula = [["=INDIRECT(\"A0\")"]];
    private static readonly string[] FormulaErrorChecks = ["formula-errors"];
    private static readonly string[] ControlTotalChecks = ["control-totals"];
    private static readonly object[][] ControlTotalValue = [[100d]];
    private static readonly Dictionary<string, object?>[] ControlTotals =
    [
        new()
        {
            ["sheetName"] = "Sheet1",
            ["cellAddress"] = "B1",
            ["expectedValue"] = 100d,
            ["tolerance"] = 0d
        }
    ];

    private readonly string _testExcelFile;
    private string? _sessionId;

    public WorkbookIntegrityProtocolTests(ITestOutputHelper output)
        : base(output, "WorkbookIntegrityProtocolClient")
    {
        _testExcelFile = Path.Join(
            CreateTempDirectory("WorkbookIntegrityProtocol"),
            "WorkbookIntegrity.xlsx");
    }

    protected override async Task InitializeTestAsync()
    {
        _sessionId = await CreateWorkbookSessionAsync(_testExcelFile);
    }

    [Fact]
    public async Task ValidateIntegrity_ReturnsTypedCanonicalFindingThroughMcp()
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

        var resultJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "validate-integrity",
            ["session_id"] = _sessionId,
            ["checks"] = FormulaErrorChecks,
            ["worksheet_names"] = """["Sheet1"]""",
            ["max_findings"] = 10
        });

        AssertIntegrityReferenceError(resultJson);

        var setValueJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = _sessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "B1",
            ["values"] = ControlTotalValue
        });
        AssertSetupSuccess(setValueJson, "range.set-values");

        var controlTotalJson = await CallToolAsync("workbook", new Dictionary<string, object?>
        {
            ["action"] = "validate-integrity",
            ["session_id"] = _sessionId,
            ["checks"] = ControlTotalChecks,
            ["control_totals"] = ControlTotals
        });
        using (var controlDocument = JsonDocument.Parse(controlTotalJson))
        {
            var root = controlDocument.RootElement;
            Assert.True(root.GetProperty("success").GetBoolean(), controlTotalJson);
            Assert.Equal("passed", root.GetProperty("overallStatus").GetString());
            Assert.Equal("control-totals", root.GetProperty("checkedChecks")[0].GetString());
        }

        await CloseSessionAsync(_sessionId, save: false);
        _sessionId = null;
    }

    private static void AssertIntegrityReferenceError(string json)
    {
        using var document = JsonDocument.Parse(json);
        var root = document.RootElement;
        Assert.True(root.GetProperty("success").GetBoolean(), json);
        Assert.Equal("failed", root.GetProperty("overallStatus").GetString());
        Assert.Equal(1, root.GetProperty("findingCount").GetInt32());
        Assert.False(root.GetProperty("findingsTruncated").GetBoolean());
        Assert.Equal("formula-errors", root.GetProperty("checkedChecks")[0].GetString());

        var group = Assert.Single(root.GetProperty("groups").EnumerateArray());
        Assert.Equal("error", group.GetProperty("severity").GetString());
        Assert.Equal("broken-reference", group.GetProperty("category").GetString());

        var finding = Assert.Single(group.GetProperty("findings").EnumerateArray());
        Assert.Equal("broken-formula-reference", finding.GetProperty("code").GetString());
        Assert.Equal("deterministic", finding.GetProperty("reliability").GetString());
        Assert.Equal("Sheet1", finding.GetProperty("sheetName").GetString());
        Assert.Equal("A1", finding.GetProperty("cellAddress").GetString());
        Assert.Equal("=INDIRECT(\"A0\")", finding.GetProperty("formula").GetString());
        Assert.Equal("#REF!", finding.GetProperty("errorName").GetString());
        Assert.Equal(-2146826265, finding.GetProperty("errorCode").GetInt32());
    }
}
