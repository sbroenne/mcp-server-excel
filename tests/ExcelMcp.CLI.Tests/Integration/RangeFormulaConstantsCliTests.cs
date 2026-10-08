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
public sealed class RangeFormulaConstantsCliTests
{
    [Theory]
    [InlineData("set-formulas")]
    [InlineData("validate-formulas")]
    public async Task MixedFormulaCells_AreAcceptedAndForwarded(string action)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", action, "--session", "session-1", "--sheet-name", "Checks",
            "--range-address", "A1:E1", "--formulas", """[["Label",5.86,true,null,"=1+1"]]"""
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });

        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"range.{action}", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        using var expected = JsonDocument.Parse("""[["Label",5.86,true,null,"=1+1"]]""");
        var formulas = args.RootElement.GetProperty("formulas");
        Assert.True(JsonElement.DeepEquals(expected.RootElement, formulas),
            $"Expected forwarded formulas {expected.RootElement.GetRawText()}, but found {formulas.GetRawText()}.");
    }
}
