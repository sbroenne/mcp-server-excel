using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Workflow")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class WorkflowToolRealExcelTests : McpIntegrationTestBase
{
    public WorkflowToolRealExcelTests(ITestOutputHelper output)
        : base(output, "WorkflowToolRealExcelClient")
    {
    }

    protected override IReadOnlyList<string> ServerArguments =>
        ["--tool-profile", "copilot-compact"];

    [Fact]
    public async Task OpenDescribeAndExecutePlan_CompleteThroughThePublicMcpTransport()
    {
        var directory = CreateTempDirectory("WorkflowToolRealExcel");
        var workbookPath = Path.Join(directory, "workflow-public.xlsx");
        var seedSessionId = await CreateWorkbookSessionAsync(workbookPath);

        var seedResult = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = seedSessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1",
            ["values"] = new List<List<object?>> { new() { "Header" } },
        });
        AssertSuccess(seedResult, "range.set-values seed");
        await CloseSessionAsync(seedSessionId, save: true);

        var openResult = await CallToolAsync("workflow", new Dictionary<string, object?>
        {
            ["action"] = "open-and-describe",
            ["file_path"] = workbookPath,
            ["preview_rows"] = 2,
            ["preview_columns"] = 2,
        });
        using var openDocument = JsonDocument.Parse(openResult);
        var openRoot = openDocument.RootElement;
        Assert.True(openRoot.GetProperty("success").GetBoolean(), openResult);
        var sessionId = openRoot.GetProperty("sessionId").GetString();
        Assert.False(string.IsNullOrWhiteSpace(sessionId));
        TrackSession(sessionId);
        Assert.Equal("Header", openRoot.GetProperty("sheets")[0].GetProperty("preview")[0][0].GetString());

        object?[] operations =
        [
            new Dictionary<string, object?>
            {
                ["command"] = "range.set-values",
                ["args"] = new Dictionary<string, object?>
                {
                    ["sheetName"] = "Sheet1",
                    ["rangeAddress"] = "A2",
                    ["values"] = new object?[][] { [10] },
                },
            },
            new Dictionary<string, object?>
            {
                ["command"] = "range.set-values",
                ["args"] = new Dictionary<string, object?>
                {
                    ["sheetName"] = "Sheet1",
                    ["rangeAddress"] = "A3",
                    ["values"] = new object?[][] { [20] },
                },
            },
        ];
        var executeResult = await CallToolAsync("workflow", new Dictionary<string, object?>
        {
            ["action"] = "execute-plan",
            ["session_id"] = sessionId,
            ["operations"] = operations,
            ["stop_on_error"] = true,
            ["verify_sheet_name"] = "Sheet1",
            ["verify_range_address"] = "A2:A3",
        });
        using var executeDocument = JsonDocument.Parse(executeResult);
        Assert.Equal("completed", executeDocument.RootElement.GetProperty("outcome").GetString());
        Assert.Equal(2, executeDocument.RootElement.GetProperty("attemptedCount").GetInt32());
        Assert.Equal(2, executeDocument.RootElement.GetProperty("completedCount").GetInt32());
        Assert.Equal(2, executeDocument.RootElement.GetProperty("steps").GetArrayLength());
        Assert.Equal("fast", executeDocument.RootElement.GetProperty("executionMode").GetString());
        Assert.True(executeDocument.RootElement.GetProperty("fastModeUsed").GetBoolean(), executeResult);
        Assert.Equal(1, executeDocument.RootElement.GetProperty("staDispatchCount").GetInt64());
        var verification = executeDocument.RootElement.GetProperty("verification");
        Assert.Equal("verified", verification.GetProperty("status").GetString());
        Assert.Equal("Sheet1", verification.GetProperty("sheetName").GetString());
        Assert.Equal("$A$2:$A$3", verification.GetProperty("rangeAddress").GetString());
        Assert.Equal(2, verification.GetProperty("rowCount").GetInt32());
        Assert.Equal(1, verification.GetProperty("columnCount").GetInt32());
        Assert.Equal(2, verification.GetProperty("cellCount").GetInt32());
        Assert.Equal(2, verification.GetProperty("nonEmptyCellCount").GetInt32());
        Assert.Equal(0, verification.GetProperty("formulaCellCount").GetInt32());
        Assert.Equal(64, verification.GetProperty("fingerprint").GetString()!.Length);
        var firstFingerprint = verification.GetProperty("fingerprint").GetString();
        var preview = verification.GetProperty("preview");
        Assert.Equal(10d, preview[0][0].GetDouble());
        Assert.Equal(20d, preview[1][0].GetDouble());

        var repeatResult = await CallToolAsync("workflow", new Dictionary<string, object?>
        {
            ["action"] = "execute-plan",
            ["session_id"] = sessionId,
            ["operations"] = new object?[]
            {
                new Dictionary<string, object?>
                {
                    ["command"] = "range.get-values",
                    ["args"] = new Dictionary<string, object?>
                    {
                        ["sheetName"] = "Sheet1",
                        ["rangeAddress"] = "A2:A3",
                    },
                },
            },
            ["verify_sheet_name"] = "Sheet1",
            ["verify_range_address"] = "A2:A3",
        });
        using var repeatDocument = JsonDocument.Parse(repeatResult);
        Assert.Equal("completed", repeatDocument.RootElement.GetProperty("outcome").GetString());
        Assert.Equal(1, repeatDocument.RootElement.GetProperty("staDispatchCount").GetInt64());
        Assert.Equal(
            firstFingerprint,
            repeatDocument.RootElement.GetProperty("verification").GetProperty("fingerprint").GetString());

        await CloseSessionAsync(sessionId, save: false);
    }

    [Fact]
    public async Task ExecutePlan_LargeVerificationScope_IsExplicitlyBoundedThroughPublicMcp()
    {
        var directory = CreateTempDirectory("WorkflowVerificationBound");
        var workbookPath = Path.Join(directory, "workflow-bounded-verification.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);

        var result = await CallToolAsync("workflow", new Dictionary<string, object?>
        {
            ["action"] = "execute-plan",
            ["session_id"] = sessionId,
            ["operations"] = new object?[]
            {
                new Dictionary<string, object?>
                {
                    ["command"] = "range.set-values",
                    ["args"] = new Dictionary<string, object?>
                    {
                        ["sheetName"] = "Sheet1",
                        ["rangeAddress"] = "A1",
                        ["values"] = new object?[][] { ["marker"] },
                    },
                },
            },
            ["verify_sheet_name"] = "Sheet1",
            ["verify_range_address"] = "A1:CV101",
        });

        using var document = JsonDocument.Parse(result);
        var root = document.RootElement;
        Assert.Equal("completed", root.GetProperty("outcome").GetString());
        Assert.Equal(1, root.GetProperty("staDispatchCount").GetInt64());
        var verification = root.GetProperty("verification");
        Assert.Equal("partiallyVerified", verification.GetProperty("status").GetString());
        Assert.Equal("$A$1:$CV$101", verification.GetProperty("rangeAddress").GetString());
        Assert.Equal(101, verification.GetProperty("rowCount").GetInt32());
        Assert.Equal(100, verification.GetProperty("columnCount").GetInt32());
        Assert.Equal(10_100, verification.GetProperty("cellCount").GetInt64());
        Assert.Equal(9_999, verification.GetProperty("inspectedCellCount").GetInt32());
        Assert.Equal("$A$1:$CU$101", verification.GetProperty("inspectedRangeAddress").GetString());
        Assert.Equal(1, verification.GetProperty("nonEmptyCellCount").GetInt32());
        Assert.Equal(0, verification.GetProperty("formulaCellCount").GetInt32());
        Assert.Equal(64, verification.GetProperty("fingerprint").GetString()!.Length);
        Assert.False(string.IsNullOrWhiteSpace(verification.GetProperty("limitation").GetString()));
        var preview = verification.GetProperty("preview");
        Assert.Equal(2, preview.GetArrayLength());
        Assert.Equal(4, preview[0].GetArrayLength());
        Assert.Equal("marker", preview[0][0].GetString());

        await CloseSessionAsync(sessionId, save: false);
    }
}
