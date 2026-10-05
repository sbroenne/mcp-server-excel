using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeMergedCellWriteRegressionTests(
    RecordingProgramTransportFixture fixture)
{
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task SetValues_MergedNonAnchorCell_ReturnsActionableFailureViaMcp()
    {
        const string sessionId = "recording-session";
        var call = await _fixture.CallToolAsync(
            "range",
            new Dictionary<string, object?>
            {
                ["action"] = "set-values",
                ["workbook_session_id"] = sessionId,
                ["sheet_name"] = "Sheet1",
                ["range_address"] = "B1",
                ["values"] = new List<List<object?>>
                {
                    new() { "Updated" }
                }
            },
            new ServiceResponse
            {
                Success = false,
                Command = "range.set-values",
                SessionId = sessionId,
                ErrorMessage =
                    "Cannot write to merged range $A$1:$B$1 outside its top-left cell; unmerge first.",
                ExceptionType = "OperationFailureException",
                ErrorCategory = "Conflict"
            },
            "range.set-values",
            """{"sheetName":"Sheet1","rangeAddress":"B1","values":[["Updated"]]}""");

        using (var args = RecordingToolTest.ParseArgs(
            call.Request,
            "range.set-values",
            sessionId))
        {
            Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
            Assert.Equal("B1", args.RootElement.GetProperty("rangeAddress").GetString());
            Assert.Equal(
                "Updated",
                args.RootElement.GetProperty("values")[0][0].GetString());
        }

        using var result = JsonDocument.Parse(call.JsonResult);
        var root = result.RootElement;
        Assert.False(root.GetProperty("success").GetBoolean());
        Assert.Equal(
            "OperationFailureException",
            root.GetProperty("exceptionType").GetString());
        Assert.Equal("Conflict", root.GetProperty("errorCategory").GetString());
        var error = root.GetProperty("errorMessage").GetString();
        Assert.Contains("$A$1:$B$1", error, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("top-left", error, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("unmerge", error, StringComparison.OrdinalIgnoreCase);
    }
}
