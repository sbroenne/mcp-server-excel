using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "NativeDataCleanup")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class NativeCleanupContractCliTests
{
    [Fact]
    public async Task RemoveDuplicates_MapsExplicitRelativeKeysAndHeaderMode()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "remove-duplicates", "--session", "session-1",
            "--sheet-name", "Data", "--range-address", "A1:C10",
            "--key-columns", "[1,3]", "--has-headers", "false"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true,"removedRows":2,"remainingRows":8}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeedit.remove-duplicates", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal([1, 3], args.RootElement.GetProperty("keyColumns").EnumerateArray().Select(value => value.GetInt32()));
        Assert.False(args.RootElement.GetProperty("hasHeaders").GetBoolean());
    }

    [Fact]
    public async Task TextToColumns_MapsTypedNestedFieldsAndOverwritePermission()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "text-to-columns", "--session", "session-1",
            "--sheet-name", "Data", "--source-range", "A1:A10", "--destination-cell", "D1",
            "--options", """{"comma":true,"fields":[{"position":1,"dataType":"Text"}],"trailingMinusNumbers":false}""",
            "--overwrite-policy", "allow"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true,"destinationRange":"$D$1:$F$10","outputColumns":3}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeedit.text-to-columns", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        var options = args.RootElement.GetProperty("options");
        Assert.True(options.GetProperty("comma").GetBoolean());
        Assert.False(options.GetProperty("trailingMinusNumbers").GetBoolean());
        Assert.Equal("Text", options.GetProperty("fields")[0].GetProperty("dataType").GetString());
        Assert.Equal("allow", args.RootElement.GetProperty("overwritePolicy").GetString(), ignoreCase: true);
    }

    [Theory]
    [InlineData("""{"comma":true,"unknown":true}""")]
    [InlineData("""{"comma":true,"fields":[{"position":1,"dataType":"wrong"}]}""")]
    public async Task TextToColumns_RejectsUnknownNestedSettingsBeforeDispatch(string options)
    {
        int calls = 0;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeedit", "text-to-columns", "--session", "session-1",
            "--sheet-name", "Data", "--source-range", "A1", "--destination-cell", "D1", "--options", options
        ], _ =>
        {
            calls++;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.NotEqual(0, result.ExitCode);
        Assert.Equal(0, calls);
    }
}
