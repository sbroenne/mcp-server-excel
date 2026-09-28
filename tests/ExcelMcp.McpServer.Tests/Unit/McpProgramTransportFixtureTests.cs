using Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class McpProgramTransportFixtureTests
{
    [Fact]
    public async Task CloseTrackedSessionsAsync_FailedResponseIsReported_AndLaterSessionStillCloses()
    {
        var trackedSessionIds = new HashSet<string>(StringComparer.Ordinal)
        {
            "failed-session",
            "successful-session"
        };
        var attemptedSessionIds = new List<string>();

        var failures = await McpProgramTransportFixture.CloseTrackedSessionsAsync(
            trackedSessionIds,
            sessionId =>
            {
                attemptedSessionIds.Add(sessionId);
                return Task.FromResult(
                    sessionId == "failed-session"
                        ? """{"success":false,"errorMessage":"close rejected"}"""
                        : """{"success":true}""");
            });

        Assert.Equal(2, attemptedSessionIds.Count);
        Assert.Contains("failed-session", attemptedSessionIds);
        Assert.Contains("successful-session", attemptedSessionIds);
        Assert.Single(failures);
        Assert.Contains("file.close failed", failures[0].Message, StringComparison.Ordinal);
        Assert.Contains("failed-session", trackedSessionIds);
        Assert.DoesNotContain("successful-session", trackedSessionIds);
    }
}
