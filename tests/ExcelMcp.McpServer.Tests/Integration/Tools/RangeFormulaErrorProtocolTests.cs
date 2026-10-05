using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Range")]
[Trait("RequiresExcel", "false")]
public sealed class RangeFormulaErrorProtocolTests(
    RecordingProgramTransportFixture fixture)
{
    private const string SessionId = "recording-session";
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task RangeReads_ReturnCanonicalFormulaErrorThroughMcp()
    {
        const string responseJson = """
            {
              "success": true,
              "values": [["#REF!"]],
              "cellErrors": [{
                "cellAddress": "A1",
                "errorName": "#REF!",
                "formula": "=INDIRECT(\"A0\")",
                "errorCode": -2146826265,
                "currentValue": -2146826265
              }]
            }
            """;

        var valuesCall = await _fixture.CallToolAsync(
            "range_read",
            RangeReadArguments("get-values"),
            RecordingToolTest.Success(responseJson),
            "range.get-values",
            """{"sheetName":"Sheet1","rangeAddress":"A1"}""");
        AssertReadRequest(valuesCall.Request, "range.get-values");
        AssertCanonicalReferenceError(valuesCall.JsonResult);

        var formulasCall = await _fixture.CallToolAsync(
            "range_read",
            RangeReadArguments("get-formulas"),
            RecordingToolTest.Success(responseJson),
            "range.get-formulas",
            """{"sheetName":"Sheet1","rangeAddress":"A1"}""");
        AssertReadRequest(formulasCall.Request, "range.get-formulas");
        AssertCanonicalReferenceError(formulasCall.JsonResult);
    }

    private static Dictionary<string, object?> RangeReadArguments(
        string action) => new()
        {
            ["action"] = action,
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1"
        };

    private static void AssertReadRequest(
        Sbroenne.ExcelMcp.Service.ServiceRequest request,
        string command)
    {
        using var args = RecordingToolTest.ParseArgs(
            request,
            command,
            SessionId);
        Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("A1", args.RootElement.GetProperty("rangeAddress").GetString());
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
