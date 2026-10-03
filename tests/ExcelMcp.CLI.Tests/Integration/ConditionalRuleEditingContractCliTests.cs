using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "ConditionalRuleEditing")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class ConditionalRuleEditingContractCliTests
{
    [Theory]
    [InlineData("update-rule")]
    [InlineData("delete-rule")]
    [InlineData("set-rule-priority")]
    public async Task SelectedRule_MapsExactSelectionAndTypedChanges(string action)
    {
        ServiceRequest? captured = null;
        List<string> arguments =
        [
            "conditionalformat", action, "--session", "session-1", "--sheet-name", "Data",
            "--rule-priority", "3", "--expected-fingerprint", "rule-fingerprint"
        ];
        if (action == "update-rule")
            arguments.AddRange(["--options", """{"formula1":"15","stopIfTrue":false}"""]);
        if (action == "set-rule-priority")
            arguments.AddRange(["--new-priority", "1"]);
        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true,"rules":[]}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"conditionalformat.{action}", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(3, args.RootElement.GetProperty("rulePriority").GetInt32());
        Assert.Equal("rule-fingerprint", args.RootElement.GetProperty("expectedFingerprint").GetString());
        if (action == "update-rule")
        {
            Assert.Equal("15", args.RootElement.GetProperty("options").GetProperty("formula1").GetString());
            Assert.False(args.RootElement.GetProperty("options").GetProperty("stopIfTrue").GetBoolean());
        }
        if (action == "set-rule-priority")
            Assert.Equal(1, args.RootElement.GetProperty("newPriority").GetInt32());
    }

    [Theory]
    [InlineData(null)]
    [InlineData("""{"wrongName":true}""")]
    public async Task MissingOrUnknownOptions_DoNotDispatch(string? options)
    {
        bool dispatched = false;
        List<string> arguments =
        [
            "conditionalformat", "update-rule", "--session", "session-1", "--sheet-name", "Data",
            "--rule-priority", "3", "--expected-fingerprint", "rule-fingerprint"
        ];
        if (options is not null)
            arguments.AddRange(["--options", options]);
        var result = await InProcessCliHelper.RunAsync(arguments, _ =>
        {
            dispatched = true;
            return new ServiceResponse { Success = true };
        });
        Assert.NotEqual(0, result.ExitCode);
        Assert.False(dispatched);
    }
}
