using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
public sealed class RecordingToolTestTests
{
    [Theory]
    [InlineData("wrong.command", "recording-session", """{"sheetName":"Sheet1"}""")]
    [InlineData("sheet.list", "wrong-session", """{"sheetName":"Sheet1"}""")]
    [InlineData("sheet.list", "recording-session", "{}")]
    public void AssertRequest_MismatchedIdentityOrArguments_Throws(
        string command,
        string sessionId,
        string args)
    {
        var request = new ServiceRequest
        {
            Command = command,
            SessionId = sessionId,
            Args = args
        };

        Assert.ThrowsAny<Exception>(() => RecordingToolTest.AssertRequest(
            request,
            "sheet.list",
            "recording-session",
            """{"sheetName":"Sheet1"}"""));
    }
}
