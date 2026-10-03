using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "TableStyles")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class TableStyleContractCliTests
{
    [Theory]
    [InlineData("list-table-styles")]
    [InlineData("get-table-style")]
    [InlineData("create-table-style")]
    [InlineData("update-table-style")]
    [InlineData("delete-table-style")]
    public async Task Actions_PreserveStyleNamesAndTypedNestedOptions(string action)
    {
        ServiceRequest? captured = null;
        var arguments = new List<string> { "-q", "workbook", action, "--session", "session-1" };
        if (action != "list-table-styles") arguments.AddRange(["--style-name", "Custom"]);
        if (action == "create-table-style") arguments.AddRange(["--source-style-name", "TableStyleMedium2"]);
        if (action == "update-table-style")
            arguments.AddRange(["--table-style-options", """{"showAsAvailableTableStyle":false,"elements":[{"elementType":"xlHeaderRow","bold":false}]}"""]);
        var result = await InProcessCliHelper.RunAsync([.. arguments], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"workbook.{action}", captured.Command);
        if (action == "list-table-styles") { Assert.Null(captured.Args); return; }
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Custom", args.RootElement.GetProperty("styleName").GetString());
        if (action == "create-table-style")
            Assert.Equal("TableStyleMedium2", args.RootElement.GetProperty("sourceStyleName").GetString());
        if (action == "update-table-style")
        {
            var options = args.RootElement.GetProperty("tableStyleOptions");
            Assert.False(options.GetProperty("showAsAvailableTableStyle").GetBoolean());
            Assert.False(options.GetProperty("elements")[0].GetProperty("bold").GetBoolean());
            Assert.Equal("xlHeaderRow", options.GetProperty("elements")[0].GetProperty("elementType").GetString());
        }
    }
}
