using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

internal static class RecordingToolTest
{
    internal static ServiceResponse Success(string result) =>
        new()
        {
            Success = true,
            Result = result
        };

    internal static JsonDocument ParseArgs(
        ServiceRequest request,
        string command,
        string sessionId)
    {
        Assert.Equal(command, request.Command);
        Assert.Equal(sessionId, request.SessionId);
        Assert.NotNull(request.Args);
        return JsonDocument.Parse(request.Args);
    }

    internal static void AssertRequest(
        ServiceRequest request,
        string command,
        string sessionId,
        string? expectedArgsJson)
    {
        Assert.Equal(command, request.Command);
        Assert.Equal(sessionId, request.SessionId);

        if (expectedArgsJson is null)
        {
            Assert.Null(request.Args);
            return;
        }

        Assert.NotNull(request.Args);
        using var expected = JsonDocument.Parse(expectedArgsJson);
        using var actual = JsonDocument.Parse(request.Args);
        Assert.True(
            JsonElement.DeepEquals(expected.RootElement, actual.RootElement),
            $"Expected Service arguments {expected.RootElement.GetRawText()}, " +
            $"but received {actual.RootElement.GetRawText()}.");
    }
}
