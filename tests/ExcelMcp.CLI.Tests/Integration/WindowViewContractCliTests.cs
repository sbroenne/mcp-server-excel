using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Window")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class WindowViewContractCliTests
{
    [Theory]
    [InlineData(null)]
    [InlineData("A1")]
    [InlineData("H30:I32")]
    public async Task SetView_MapsInputsAndReturnsObservedView(string? select)
    {
        ServiceRequest? captured = null;
        List<string> command =
        [
            "-q", "window", "set-view", "--session", "session-1",
            "--sheet-name", "Dashboard", "--top-left-cell", "E20"
        ];
        if (select is not null)
            command.AddRange(["--select", select]);
        var result = await InProcessCliHelper.RunAsync(command.ToArray(), request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = true,
                Result = """{"success":true,"topLeftCell":"$E$20","selectionKind":"range","selectedRangeAddress":"$H$30:$I$32","activeCellAddress":"$H$30"}"""
            };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("window.set-view", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Dashboard", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("E20", args.RootElement.GetProperty("topLeftCell").GetString());
        Assert.Equal(select, args.RootElement.TryGetProperty("select", out var selection) ? selection.GetString() : null);
        using var output = JsonDocument.Parse(result.Stdout);
        var data = output.RootElement.GetProperty("data");
        Assert.Equal("$E$20", data.GetProperty("topLeftCell").GetString());
        Assert.Equal("range", data.GetProperty("selectionKind").GetString());
        Assert.Equal("$H$30:$I$32", data.GetProperty("selectedRangeAddress").GetString());
        Assert.Equal("$H$30", data.GetProperty("activeCellAddress").GetString());
    }

    [Theory]
    [InlineData("--top-left-cell", "A1")]
    [InlineData("--sheet-name", "Dashboard")]
    public async Task SetView_MissingRequiredInput_DoesNotDispatch(string flag, string value)
    {
        var dispatched = false;
        var result = await InProcessCliHelper.RunAsync(
            ["-q", "window", "set-view", "--session", "session-1", flag, value], _ =>
            {
                dispatched = true;
                return new ServiceResponse { Success = true, Result = """{"success":true}""" };
            });
        Assert.NotEqual(0, result.ExitCode);
        Assert.False(dispatched);
    }

    [Fact]
    public async Task SetView_BatchJson_PreservesInputsAndNativeFailure()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(["-q", "batch", "--session", "session-1"], request =>
        {
            captured = request;
            return new ServiceResponse
            {
                Success = false,
                ErrorMessage = "Excel refused the selection; the worksheet remains active.",
                ErrorCategory = "COM",
                HResult = "0x800A03EC",
                InnerError = "Native selection failure"
            };
        }, """{"command":"window.set-view","args":{"sheetName":"Dashboard","topLeftCell":"A1","select":"A1:B3"}}""");
        Assert.Equal(1, result.ExitCode);
        Assert.NotNull(captured);
        Assert.Equal("window.set-view", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal("Dashboard", args.RootElement.GetProperty("sheetName").GetString());
        Assert.Equal("A1", args.RootElement.GetProperty("topLeftCell").GetString());
        Assert.Equal("A1:B3", args.RootElement.GetProperty("select").GetString());
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.False(output.RootElement.GetProperty("success").GetBoolean());
        var error = output.RootElement.GetProperty("error");
        Assert.Equal("0x800A03EC", error.GetProperty("hresult").GetString());
        Assert.Equal("Native selection failure", error.GetProperty("innerError").GetString());
        Assert.Contains("worksheet remains active", error.GetProperty("message").GetString(), StringComparison.Ordinal);
    }
}
