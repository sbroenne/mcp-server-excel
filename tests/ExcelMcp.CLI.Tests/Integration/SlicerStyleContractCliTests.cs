using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Slicer")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SlicerStyleContractCliTests
{
    [Theory]
    [InlineData("create-style")]
    [InlineData("apply-style")]
    [InlineData("list-styles")]
    public async Task Styles_PreserveNamesAndNestedFormatting(string action)
    {
        ServiceRequest? captured = null;
        var arguments = new List<string> { "-q", "slicer", action, "--session", "session-1" };
        if (action != "list-styles") arguments.AddRange(["--style-name", "Brand"]);
        if (action == "apply-style") arguments.AddRange(["--slicer-name", "Regions"]);
        if (action == "create-style")
            arguments.AddRange(["--source-style-name", "SlicerStyleLight1", "--slicer-style-options",
                """{"selectedItemWithData":{"fillColor":"#123456","borderColor":"#654321","fontColor":"#FFFFFF","bold":false}}"""]);
        var result = await InProcessCliHelper.RunAsync([.. arguments], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"slicer.{action}", captured.Command);
        if (action == "list-styles")
        {
            Assert.Null(captured.Args);
            return;
        }
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Brand", args.RootElement.GetProperty("styleName").GetString());
        if (action == "apply-style")
        {
            Assert.Equal("Regions", args.RootElement.GetProperty("slicerName").GetString());
            Assert.Equal(2, args.RootElement.EnumerateObject().Count());
        }
        else
        {
            Assert.Equal("SlicerStyleLight1", args.RootElement.GetProperty("sourceStyleName").GetString());
            var format = args.RootElement.GetProperty("slicerStyleOptions").GetProperty("selectedItemWithData");
            Assert.Equal("#123456", format.GetProperty("fillColor").GetString());
            Assert.Equal("#654321", format.GetProperty("borderColor").GetString());
            Assert.Equal("#FFFFFF", format.GetProperty("fontColor").GetString());
            Assert.False(format.GetProperty("bold").GetBoolean());
            Assert.False(args.RootElement.GetProperty("slicerStyleOptions").TryGetProperty("header", out _));
        }
    }
}
