using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "WorkbookTheme")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class WorkbookThemeContractCliTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ThemeActions_ForwardSelectionWithoutCreatingFiles(bool apply)
    {
        ServiceRequest? captured = null;
        var arguments = new List<string> { "-q", "workbook", apply ? "apply-theme" : "get-theme", "--session", "session-1" };
        if (apply)
            arguments.AddRange(["--theme-path", "selected.thmx"]);
        var result = await InProcessCliHelper.RunAsync([.. arguments], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal(apply ? "workbook.apply-theme" : "workbook.get-theme", captured.Command);
        if (apply)
        {
            using var args = JsonDocument.Parse(captured.Args!);
            Assert.Equal("selected.thmx", args.RootElement.GetProperty("themePath").GetString());
        }
        else
            Assert.Null(captured.Args);
    }
}
