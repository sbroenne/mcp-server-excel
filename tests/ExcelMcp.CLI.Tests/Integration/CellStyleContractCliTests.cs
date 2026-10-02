using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "CellStyles")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class CellStyleContractCliTests
{
    [Theory]
    [InlineData("list-cell-styles")]
    [InlineData("get-cell-style")]
    [InlineData("create-cell-style")]
    [InlineData("update-cell-style")]
    [InlineData("delete-cell-style")]
    public async Task StyleActions_ForwardNativeSelectionAndTypedOptions(string action)
    {
        ServiceRequest? captured = null;
        var arguments = new List<string> { "-q", "workbook", action, "--session", "session-1" };
        if (action != "list-cell-styles") arguments.AddRange(["--style-name", "Custom"]);
        if (action == "create-cell-style")
            arguments.AddRange(["--source-sheet-name", "Data", "--source-cell-address", "A1"]);
        if (action == "update-cell-style")
            arguments.AddRange(["--style-options", """{"includeFont":false,"formatOptions":{"bold":false,"fillThemeColor":5}}"""]);
        var result = await InProcessCliHelper.RunAsync([.. arguments], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"workbook.{action}", captured.Command);
        if (action == "list-cell-styles")
        {
            Assert.Null(captured.Args);
            return;
        }
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Custom", args.RootElement.GetProperty("styleName").GetString());
        if (action == "create-cell-style")
        {
            Assert.Equal("Data", args.RootElement.GetProperty("sourceSheetName").GetString());
            Assert.Equal("A1", args.RootElement.GetProperty("sourceCellAddress").GetString());
        }
        if (action == "update-cell-style")
        {
            var options = args.RootElement.GetProperty("styleOptions");
            Assert.False(options.GetProperty("includeFont").GetBoolean());
            Assert.False(options.GetProperty("formatOptions").GetProperty("bold").GetBoolean());
            Assert.Equal(5, options.GetProperty("formatOptions").GetProperty("fillThemeColor").GetInt32());
        }
    }
}
