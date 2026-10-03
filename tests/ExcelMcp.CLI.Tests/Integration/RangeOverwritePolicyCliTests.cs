using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Range")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class RangeOverwritePolicyCliTests
{
    [Theory]
    [InlineData("set-values", null)]
    [InlineData("set-values", "reject-nonempty")]
    [InlineData("set-values", "allow")]
    [InlineData("set-formulas", null)]
    [InlineData("set-formulas", "reject-nonempty")]
    [InlineData("set-formulas", "allow")]
    [InlineData("copy", null)]
    [InlineData("copy", "reject-nonempty")]
    [InlineData("copy", "allow")]
    [InlineData("copy", null, "values")]
    [InlineData("copy", "reject-nonempty", "values")]
    [InlineData("copy", "allow", "values")]
    [InlineData("copy", null, "formulas")]
    [InlineData("copy", "reject-nonempty", "formulas")]
    [InlineData("copy", "allow", "formulas")]
    public async Task ContentAction_MapsOptionalPolicy(string action, string? policy, string pasteKind = "all")
    {
        var arguments = Arguments(action, pasteKind);
        if (policy is not null)
        {
            arguments.Add("--overwrite-policy");
            arguments.Add(policy);
        }
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(arguments, request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"range.{action}", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        if (policy is null)
            Assert.False(args.RootElement.TryGetProperty("overwritePolicy", out _));
        else
            Assert.Equal(policy, args.RootElement.GetProperty("overwritePolicy").GetString());
    }

    [Theory]
    [InlineData("Conflict", "Cannot write to occupied cells on sheet 'Sheet1'. Conflicting addresses: $A$1.")]
    [InlineData("InvalidInput", "Invalid value 'unknown' for parameter 'overwritePolicy'.")]
    public async Task ServiceFailure_PreservesErrorAndExitCode(string category, string message)
    {
        var result = await InProcessCliHelper.RunAsync(Arguments("set-values"), _ => new ServiceResponse
        {
            Success = false,
            ErrorCategory = category,
            ErrorMessage = message,
            ExceptionType = "OperationFailureException"
        });
        Assert.Equal(1, result.ExitCode);
        using var json = JsonDocument.Parse(result.Stdout);
        Assert.False(json.RootElement.GetProperty("success").GetBoolean());
        Assert.Equal(category, json.RootElement.GetProperty("errorCategory").GetString());
        Assert.Equal(message, json.RootElement.GetProperty("errorMessage").GetString());
    }

    [Fact]
    public async Task ReadAction_RejectsInapplicablePolicy()
    {
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", "get-values", "--session", "session-1", "--sheet-name", "Sheet1",
            "--range-address", "A1", "--overwrite-policy", "allow"
        ]);
        Assert.Equal(1, result.ExitCode);
        Assert.Contains("overwritePolicy", result.Stdout + result.Stderr, StringComparison.Ordinal);
    }

    [Fact]
    public async Task Help_AdvertisesPolicyAndProtectedDefault()
    {
        var result = await InProcessCliHelper.RunAsync(["range", "set-values", "--help"]);
        Assert.Equal(0, result.ExitCode);
        Assert.Contains("--overwrite-policy", result.Stdout, StringComparison.Ordinal);
        Assert.Contains("reject-nonempty", result.Stdout, StringComparison.Ordinal);
    }

    private static List<string> Arguments(string action, string pasteKind = "all")
    {
        List<string> arguments = ["range", action, "--session", "session-1"];
        if (action.StartsWith("copy", StringComparison.Ordinal))
            arguments.AddRange(["--source-sheet", "Sheet1", "--source-range", "A1:B2", "--target-sheet", "Sheet1", "--target-range", "D1", "--paste-kind", pasteKind]);
        else
        {
            arguments.AddRange(["--sheet-name", "Sheet1", "--range-address", "A1"]);
            arguments.AddRange(action == "set-values" ? ["--values", "[[1]]"] : ["--formulas", """[["=1"]]"""]);
        }
        return arguments;
    }
}
