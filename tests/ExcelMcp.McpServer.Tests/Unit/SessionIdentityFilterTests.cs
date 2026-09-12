using System.Text.Json;
using Sbroenne.ExcelMcp.McpServer.Telemetry;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionBinding")]
[Trait("Speed", "Fast")]
public sealed class SessionIdentityFilterTests
{
    [Fact]
    public void NormalizeSessionIdentity_AliasOnlyCopiesCanonicalAndObservesOnce()
    {
        var arguments = Arguments(("sessionId", "\"synthetic-private-value\""));
        var observations = new List<(string Tool, string Action)>();

        var error = SessionIdentityFilter.NormalizeSessionIdentity(
            arguments,
            "workbook",
            "get-info",
            (tool, action) => observations.Add((tool, action)));

        Assert.Null(error);
        Assert.Equal("synthetic-private-value", arguments["session_id"].GetString());
        Assert.Equal([("workbook", "get-info")], observations);
    }

    [Fact]
    public void NormalizeSessionIdentity_BothEqualObservesOnceWithoutReplacingCanonical()
    {
        var arguments = Arguments(
            ("session_id", "\"synthetic-private-value\""),
            ("sessionId", "\"synthetic-private-value\""));
        var observations = 0;

        var error = SessionIdentityFilter.NormalizeSessionIdentity(
            arguments,
            "file",
            "close",
            (_, _) => observations++);

        Assert.Null(error);
        Assert.Equal("synthetic-private-value", arguments["session_id"].GetString());
        Assert.Equal(1, observations);
    }

    [Fact]
    public void SessionIdAliasTelemetry_UsesOnlyFixedPrivacySafeLabels()
    {
        var telemetry = ExcelMcpTelemetry.CreateSessionIdAliasTelemetry("workbook", "get-info");

        Assert.Equal("SessionIdCompatibilityAliasObserved", telemetry.Name);
        Assert.Equal("workbook", telemetry.Properties["Tool"]);
        Assert.Equal("get-info", telemetry.Properties["Action"]);
        Assert.Equal("sessionId", telemetry.Properties["Alias"]);
        Assert.False(string.IsNullOrWhiteSpace(telemetry.Properties["AppVersion"]));
        Assert.Equal(4, telemetry.Properties.Count);
        Assert.Equal(ExcelMcpTelemetry.UserId, telemetry.Context.User.Id);
        Assert.Equal(ExcelMcpTelemetry.SessionId, telemetry.Context.Session.Id);
        Assert.Equal("ExcelMcp.McpServer", telemetry.Context.Cloud.RoleName);
        Assert.Equal($"instance-{ExcelMcpTelemetry.UserId[..8]}", telemetry.Context.Cloud.RoleInstance);
        Assert.Equal(telemetry.Properties["AppVersion"], telemetry.Context.Component.Version);

        var serialized = string.Join(
            "\n",
            telemetry.Properties.Select(property => $"{property.Key}={property.Value}"));
        Assert.DoesNotContain("synthetic-private-value", serialized, StringComparison.Ordinal);
        Assert.DoesNotContain("session_id", serialized, StringComparison.Ordinal);
        Assert.DoesNotContain("Arguments", telemetry.Properties.Keys);
        Assert.DoesNotContain("WorkbookPath", telemetry.Properties.Keys);
    }

    [Fact]
    public void SessionIdAliasTelemetry_TransportFailureDoesNotEscape()
    {
        var exception = Record.Exception(() =>
            ExcelMcpTelemetry.TryTrackSessionIdAliasObserved(
                _ => throw new InvalidOperationException("synthetic telemetry failure"),
                "workbook",
                "get-info"));

        Assert.Null(exception);
    }

    private static Dictionary<string, JsonElement> Arguments(
        params (string Name, string Json)[] values) =>
        values.ToDictionary(
            value => value.Name,
            value => JsonSerializer.Deserialize<JsonElement>(value.Json),
            StringComparer.Ordinal);
}
