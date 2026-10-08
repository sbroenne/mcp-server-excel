using System.Data.Common;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Category", "Unit")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Connection")]
[Trait("RequiresExcel", "false")]
public sealed class ConnectionAccountSettingsParsingTests
{
    [Theory]
    [InlineData("Provider=MSOLAP")]
    [InlineData("OLEDB;Provider=MSOLAP.8")]
    [InlineData("oledb;provider=msolap.7")]
    public void Parse_AcceptsMsolapAndPreservesQuotedValues(string prefix)
    {
        var settings = ConnectionCommands.ParseAccountSettingsConnectionString(
            prefix + ";User ID=\"fixture;account\";Password=\"fixture;secret\";EffectiveUserName=fixture-impersonation;Custom Setting=\"unchanged;value\";");
        Assert.Equal(5, settings.Count);
        Assert.Equal("fixture;account", settings["User ID"]);
        Assert.Equal("fixture;secret", settings["Password"]);
        Assert.Equal("fixture-impersonation", settings["EffectiveUserName"]);
        Assert.Equal("unchanged;value", settings["Custom Setting"]);
    }

    [Theory]
    [InlineData("")]
    [InlineData("Provider=MSOLAP.")]
    [InlineData("Provider=MSOLAP.8.other")]
    [InlineData("Provider=MSOLAPOther")]
    [InlineData("Provider=Microsoft.Mashup.OleDb.1")]
    [InlineData("Provider=Microsoft.ACE.OLEDB.12.0")]
    [InlineData("Provider=private-provider-value")]
    public void Parse_RejectsOtherProvidersWithoutExposingValues(string settings)
    {
        var error = Assert.Throws<NotSupportedException>(() =>
            ConnectionCommands.ParseAccountSettingsConnectionString(settings + ";User ID=fixture-account;Password=fixture-secret;"));
        Assert.Contains("MSOLAP", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("fixture-account", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("fixture-secret", error.Message, StringComparison.Ordinal);
        Assert.DoesNotContain("private-provider-value", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(null, null)]
    [InlineData("always", "Always")]
    [InlineData("", "Unrecognized")]
    [InlineData("private-mode-value", "Unrecognized")]
    public void SignInModes_DoNotGuessOrExposeUnrecognizedValues(string? mode, string? expected)
    {
        var settings = new DbConnectionStringBuilder();
        if (mode != null) settings["Interactive Login"] = mode;
        Assert.Equal(expected, ConnectionCommands.ReadSignInMode(
            settings, "Interactive Login", ["Default", "Enabled", "Disabled", "Always"]));
    }
}
