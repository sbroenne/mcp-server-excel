using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Unit;

[Trait("Layer", "McpServer")]
[Trait("Category", "Unit")]
[Trait("Feature", "SessionBinding")]
[Trait("Speed", "Fast")]
public sealed class SessionIdentityFilterTests
{
    [Fact]
    public void ValidateSessionIdentity_AliasOnlyRejectsWithoutMutatingArguments()
    {
        var arguments = Arguments(("sessionId", "\"synthetic-private-value\""));
        var error = SessionIdentityFilter.ValidateSessionIdentity(arguments);

        Assert.Equal(SessionIdentityFilter.ErrorMessage, error);
        Assert.False(arguments.ContainsKey("session_id"));
    }

    [Theory]
    [InlineData("\"synthetic-private-value\"")]
    [InlineData("\"different-private-value\"")]
    [InlineData("null")]
    [InlineData("42")]
    public void ValidateSessionIdentity_UsesOnlyCanonicalIdentity(string extraValue)
    {
        var arguments = Arguments(
            ("session_id", "\"synthetic-private-value\""),
            ("sessionId", extraValue));
        var error = SessionIdentityFilter.ValidateSessionIdentity(arguments);

        Assert.Null(error);
        Assert.Equal("synthetic-private-value", arguments["session_id"].GetString());
    }

    [Theory]
    [InlineData("null")]
    [InlineData("\"\"")]
    [InlineData("\"   \"")]
    [InlineData("42")]
    [InlineData("true")]
    [InlineData("{}")]
    [InlineData("[]")]
    public void ValidateSessionIdentity_InvalidCanonicalReturnsRequiredInputError(string value)
    {
        var arguments = Arguments(("session_id", value));

        Assert.Equal(SessionIdentityFilter.ErrorMessage,
            SessionIdentityFilter.ValidateSessionIdentity(arguments));
    }

    private static Dictionary<string, JsonElement> Arguments(
        params (string Name, string Json)[] values) =>
        values.ToDictionary(
            value => value.Name,
            value => JsonSerializer.Deserialize<JsonElement>(value.Json),
            StringComparer.Ordinal);
}
