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
public sealed class RangeTraceContractCliTests
{
    [Theory]
    [InlineData("trace-precedents", true)]
    [InlineData("trace-precedents", false)]
    [InlineData("trace-dependents", true)]
    [InlineData("trace-dependents", false)]
    public async Task Trace_MapsLimitsAndPreservesPartialResults(string action, bool useDefaults)
    {
        ServiceRequest? captured = null;
        const string response =
            """{"success":true,"coverage":{"scope":"same-worksheet-only","workbookComplete":false,"nativeTraversalComplete":false},"limits":{"maxDepth":1,"maxNodes":2,"maxPreviewCharacters":8},"truncation":{"any":true,"depthLimitReached":true,"nodeLimitReached":false,"startingCellsOmitted":false,"previewTextTruncated":true},"nodes":[{"address":"$A$1","depth":0,"formula":"=SUM(B1:","formulaTruncated":true}],"edges":[],"cycles":[],"unresolved":[],"boundaries":[{"address":"$A$1","reason":"max-depth"}]}""";
        List<string> arguments =
        [
            "range", action, "--session", "session-1", "--sheet-name", "Sheet1",
            "--range-address", "A1,C3"
        ];
        if (!useDefaults)
            arguments.AddRange(["--max-depth", "1", "--max-nodes", "2", "--max-preview-characters", "8"]);
        var result = await InProcessCliHelper.RunAsync(arguments.ToArray(), request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = true,
                Result = response
            };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"range.{action}", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("A1,C3", args.RootElement.GetProperty("rangeAddress").GetString());
        if (useDefaults)
        {
            Assert.False(args.RootElement.TryGetProperty("maxDepth", out _));
            Assert.False(args.RootElement.TryGetProperty("maxNodes", out _));
            Assert.False(args.RootElement.TryGetProperty("maxPreviewCharacters", out _));
        }
        else
        {
            Assert.Equal(1, args.RootElement.GetProperty("maxDepth").GetInt32());
            Assert.Equal(2, args.RootElement.GetProperty("maxNodes").GetInt32());
            Assert.Equal(8, args.RootElement.GetProperty("maxPreviewCharacters").GetInt32());
        }
        using var output = JsonDocument.Parse(result.Stdout);
        using var expected = JsonDocument.Parse(response);
        var data = output.RootElement.GetProperty("data");
        Assert.Equal(expected.RootElement.EnumerateObject().Count() - 1, data.EnumerateObject().Count());
        foreach (var property in expected.RootElement.EnumerateObject().Where(property => property.Name != "success"))
            Assert.True(JsonElement.DeepEquals(property.Value, data.GetProperty(property.Name)), property.Name);
    }

    [Theory]
    [InlineData("trace-precedents")]
    [InlineData("trace-dependents")]
    public async Task Trace_InvalidLimitTypeDoesNotDispatch(string action)
    {
        bool dispatched = false;
        var result = await InProcessCliHelper.RunAsync(
        [
            "range", action, "--session", "session-1", "--sheet-name", "Sheet1",
            "--range-address", "A1", "--max-depth", "not-a-number"
        ], _ =>
        {
            dispatched = true;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.NotEqual(0, result.ExitCode);
        Assert.False(dispatched);
        Assert.Contains("not-a-number", result.Stdout + result.Stderr, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("trace-precedents")]
    [InlineData("trace-dependents")]
    public async Task Trace_BatchJsonPreservesLimits(string action)
    {
        const string args = """{"sheetName":"Sheet1","rangeAddress":"A1,C3","maxDepth":0,"maxNodes":2,"maxPreviewCharacters":8}""";
        string input = $$"""{"command":"range.{{action}}","sessionId":"session-1","args":{{args}}}""";
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(["batch", "--input", "-"], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        }, input + Environment.NewLine);
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal($"range.{action}", captured.Command);
        using var actual = JsonDocument.Parse(captured.Args!);
        using var expected = JsonDocument.Parse(args);
        Assert.True(JsonElement.DeepEquals(expected.RootElement, actual.RootElement));
    }
}
