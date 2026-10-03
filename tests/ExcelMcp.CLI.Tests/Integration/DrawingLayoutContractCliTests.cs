using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "DrawingLayout")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class DrawingLayoutContractCliTests
{
    private static readonly string[] ObjectNames = ["First", "Second", "Third"];

    [Theory]
    [InlineData("group-objects", "--group-name", "Together", "groupName")]
    [InlineData("align-objects", "--alignment", "Left", "alignment")]
    [InlineData("distribute-objects", "--distribution", "Horizontal", "distribution")]
    [InlineData("duplicate-object", "--new-name", "Copy", "newName")]
    [InlineData("set-z-order", "--z-order", "BringToFront", "zOrder")]
    [InlineData("ungroup-object", "--object-name", "Together", "objectName")]
    public async Task Layout_MapsNamesArraysEnumsAndDefaults(string action, string flag, string value, string property)
    {
        ServiceRequest? captured = null;
        var multiple = action is "group-objects" or "align-objects" or "distribute-objects";
        var args = new List<string> { "-q", "drawing", action, "--session", "session-1", "--sheet-name", "Dashboard" };
        if (multiple)
            args.AddRange(["--object-names", """["First","Second","Third"]"""]);
        else if (action != "ungroup-object")
            args.AddRange(["--object-name", "First"]);
        args.AddRange([flag, value]);
        var result = await InProcessCliHelper.RunAsync([.. args], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true,"drawingObjects":[{"name":"Actual"}]}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"drawing.{action}", captured.Command);
        using var document = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Dashboard", document.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal(value, document.RootElement.GetProperty(property).GetString());
        if (multiple)
            Assert.Equal(ObjectNames, document.RootElement.GetProperty("objectNames").EnumerateArray().Select(item => item.GetString()));
        if (action == "duplicate-object")
        {
            Assert.False(document.RootElement.TryGetProperty("offsetLeft", out _));
            Assert.False(document.RootElement.TryGetProperty("offsetTop", out _));
        }
    }
}
