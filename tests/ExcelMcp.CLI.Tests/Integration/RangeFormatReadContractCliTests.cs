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
public sealed class RangeFormatReadContractCliTests
{
    [Theory]
    [InlineData("stored")]
    [InlineData("displayed")]
    [InlineData("both")]
    [InlineData(null)]
    public async Task GetFormat_MapsViewAndPreservesEveryCell(string? view)
    {
        ServiceRequest? captured = null;
        var response = JsonSerializer.Serialize(new
        {
            success = true,
            cellCount = 64,
            cells = Enumerable.Range(1, 64).Select(row => new { address = $"$A${row}" }).ToArray()
        });
        List<string> arguments =
        [
            "rangeformat", "get-format", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1:A64"
        ];
        if (view is not null)
        {
            arguments.AddRange(["--view", view]);
        }
        var result = await InProcessCliHelper.RunAsync(arguments.ToArray(), request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = response };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeformat.get-format", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Sheet1", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("A1:A64", args.RootElement.GetProperty("rangeAddress").GetString());
        if (view is null)
        {
            Assert.False(args.RootElement.TryGetProperty("view", out _));
        }
        else
        {
            Assert.Equal(view, args.RootElement.GetProperty("view").GetString(), ignoreCase: true);
        }
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal(64, output.RootElement.GetProperty("cellCount").GetInt64());
        Assert.Equal(64, output.RootElement.GetProperty("cells").GetArrayLength());
    }

    [Theory]
    [InlineData("unknown")]
    [InlineData("99")]
    public async Task GetFormat_InvalidViewDoesNotDispatch(string view)
    {
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeformat", "get-format", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1", "--view", view
        ]);
        Assert.NotEqual(0, result.ExitCode);
        Assert.Contains("view", result.Stdout + result.Stderr, StringComparison.OrdinalIgnoreCase);
    }
}
