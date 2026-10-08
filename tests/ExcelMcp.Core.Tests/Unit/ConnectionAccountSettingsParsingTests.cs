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
    [InlineData(null, null, null)]
    [InlineData("", null, null)]
    [InlineData(" ", 0, null)]
    [InlineData("fixture\0account", null, null)]
    [InlineData("fixture-account", 999, null)]
    [InlineData("fixture-account", null, 999)]
    public void UpdateValidation_RejectsInvalidInputsWithoutEchoingAccount(
        string? hint, int? interactive, int? identity)
    {
        var error = Assert.Throws<ArgumentException>(() =>
            ConnectionCommands.ValidateAccountSettingsUpdate(hint,
                interactive.HasValue ? (ConnectionInteractiveLogin)interactive.Value : null,
                identity.HasValue ? (ConnectionIdentityMode)identity.Value : null));
        Assert.DoesNotContain("fixture", error.Message, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(0, 0)]
    [InlineData(1, 1)]
    [InlineData(2, 2)]
    [InlineData(3, 3)]
    public void UpdateValidation_AcceptsEveryExplicitMode(int interactive, int identity) =>
        ConnectionCommands.ValidateAccountSettingsUpdate(
            null, (ConnectionInteractiveLogin)interactive, (ConnectionIdentityMode)identity);

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
