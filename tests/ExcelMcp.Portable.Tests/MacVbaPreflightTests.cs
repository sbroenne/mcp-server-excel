using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacVbaPreflightTests
{
    [Fact]
    public void EntirelyDisabled_OverridesOtherMacroPreferences()
    {
        var result = MacVbaPreflight.Evaluate(new Dictionary<string, string?>
        {
            ["VisualBasicEntirelyDisabled"] = "1",
            ["VisualBasicMacroExecutionState"] = "EnabledWithoutWarnings",
            ["VBAObjectModelIsTrusted"] = "1"
        });

        Assert.Equal(MacMacroExecutionAvailability.Disabled, result.MacroExecution);
        Assert.Equal(MacVbaProjectModelAccess.Enabled, result.ProjectModelAccess);
    }

    [Theory]
    [InlineData("EnabledWithoutWarnings", "Available")]
    [InlineData("DisabledWithoutWarnings", "Disabled")]
    [InlineData("DisabledWithWarnings", "UserApprovalRequired")]
    [InlineData(null, "UserApprovalRequired")]
    public void MacroExecutionPreference_IsClassifiedWithoutPrompting(
        string? setting,
        string expectedName)
    {
        var expected = Enum.Parse<MacMacroExecutionAvailability>(expectedName);
        var preferences = new Dictionary<string, string?>();
        if (setting is not null)
        {
            preferences["VisualBasicMacroExecutionState"] = setting;
        }

        var result = MacVbaPreflight.Evaluate(preferences);

        Assert.Equal(expected, result.MacroExecution);
    }

    [Theory]
    [InlineData("1", "Enabled")]
    [InlineData("true", "Enabled")]
    [InlineData("0", "Disabled")]
    [InlineData(null, "Disabled")]
    public void ProjectModelTrust_IsReadOnlyAndExplicit(
        string? setting,
        string expectedName)
    {
        var expected = Enum.Parse<MacVbaProjectModelAccess>(expectedName);
        var preferences = new Dictionary<string, string?>();
        if (setting is not null)
        {
            preferences["VBAObjectModelIsTrusted"] = setting;
        }

        var result = MacVbaPreflight.Evaluate(preferences);

        Assert.Equal(expected, result.ProjectModelAccess);
    }

    [Fact]
    public void UnknownMacroPreference_DoesNotAssumeAvailability()
    {
        var result = MacVbaPreflight.Evaluate(new Dictionary<string, string?>
        {
            ["VisualBasicMacroExecutionState"] = "UnexpectedFutureValue"
        });

        Assert.Equal(MacMacroExecutionAvailability.Unknown, result.MacroExecution);
    }
}
