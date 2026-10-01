using System.Text.Json.Nodes;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacRangeArgumentsTests : IDisposable
{
    private readonly string _directory = Directory.CreateTempSubdirectory("excelmcp-mac-range-args-").FullName;

    [Fact]
    public void Prepare_SetNumberFormats_PreservesInlineMatrixAndRemovesFileArgument()
    {
        var arguments = new JsonObject
        {
            ["formats"] = new JsonArray
            {
                new JsonArray("#,##0.00", "0.00%"),
                new JsonArray("m/d/yyyy", "General")
            },
            ["formatsFile"] = null
        };

        MacRangeArguments.Prepare("range", "set-number-formats", arguments);

        Assert.Equal("#,##0.00", arguments["formats"]![0]![0]!.GetValue<string>());
        Assert.Equal("General", arguments["formats"]![1]![1]!.GetValue<string>());
        Assert.False(arguments.ContainsKey("formatsFile"));
    }

    [Fact]
    public void Prepare_SetNumberFormats_LoadsJsonFile()
    {
        var path = Path.Combine(_directory, "formats.json");
        File.WriteAllText(path, """[["0.00","@"],["General","m/d/yyyy"]]""");
        var arguments = new JsonObject { ["formatsFile"] = path };

        MacRangeArguments.Prepare("range", "set-number-formats", arguments);

        Assert.Equal("0.00", arguments["formats"]![0]![0]!.GetValue<string>());
        Assert.Equal("m/d/yyyy", arguments["formats"]![1]![1]!.GetValue<string>());
        Assert.False(arguments.ContainsKey("formatsFile"));
    }

    [Fact]
    public void Prepare_SetNumberFormats_RequiresInlineOrFileInput()
    {
        var error = Assert.Throws<ArgumentException>(
            () => MacRangeArguments.Prepare("range", "set-number-formats", new JsonObject()));

        Assert.Contains("formats", error.Message, StringComparison.Ordinal);
    }

    public void Dispose() => Directory.Delete(_directory, recursive: true);
}
