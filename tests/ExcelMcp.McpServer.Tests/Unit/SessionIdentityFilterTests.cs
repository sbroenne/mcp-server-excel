using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionBinding")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class SessionIdentityFilterTests
{
    [Theory]
    [InlineData("null")]
    [InlineData("\"\"")]
    [InlineData("\"   \"")]
    [InlineData("42")]
    [InlineData("true")]
    [InlineData("{}")]
    [InlineData("[]")]
    public void ValidateSessionIdentity_RejectsInvalidCanonicalValue(string json)
    {
        var arguments = Arguments(("session_id", json));

        var error = SessionIdentityFilter.ValidateSessionIdentity(arguments);

        Assert.Equal(SessionIdentityFilter.ErrorMessage, error);
    }

    [Fact]
    public void ValidateSessionIdentity_AcceptsCanonicalWithoutMutation()
    {
        var arguments = Arguments(("session_id", "\"synthetic-private-value\""));

        var error = SessionIdentityFilter.ValidateSessionIdentity(arguments);

        Assert.Null(error);
        Assert.Equal("synthetic-private-value", arguments["session_id"].GetString());
        Assert.Single(arguments);
    }

    [Fact]
    public void ValidateSessionIdentity_MissingCanonicalIsNotFilledFromLegacyName()
    {
        var arguments = Arguments(("sessionId", "\"synthetic-private-value\""));

        var error = SessionIdentityFilter.ValidateSessionIdentity(arguments);

        Assert.Equal(SessionIdentityFilter.ErrorMessage, error);
        Assert.False(arguments.ContainsKey("session_id"));
        Assert.Single(arguments);
    }

    [Fact]
    public void ValidateSessionIdentity_MissingIdentityReturnsRecoveryGuidance()
    {
        var error = SessionIdentityFilter.ValidateSessionIdentity(Arguments());

        Assert.Equal(SessionIdentityFilter.ErrorMessage, error);
        Assert.Contains("session_id", error, StringComparison.Ordinal);
        Assert.Contains("file list", error, StringComparison.Ordinal);
    }

    private static Dictionary<string, JsonElement> Arguments(
        params (string Name, string Json)[] values) =>
        values.ToDictionary(
            value => value.Name,
            value => JsonSerializer.Deserialize<JsonElement>(value.Json),
            StringComparer.Ordinal);
}
