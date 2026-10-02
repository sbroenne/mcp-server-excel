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
public sealed class OwnedContextVisibilityContractCliTests
{
    [Fact]
    public async Task Context_UsesOwnedSessionAndReturnsExplicitAvailability()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
            ["window", "get-context", "--session", "session-1"], request =>
            {
                captured = request;
                return new ServiceResponse
                {
                    Success = true,
                    Result = """{"success":true,"availability":"available","windows":[]}"""
                };
            });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("window.get-context", captured.Command);
        Assert.Equal("session-1", captured.SessionId);
        using var output = JsonDocument.Parse(result.Stdout);
        Assert.Equal("available", output.RootElement.GetProperty("availability").GetString());
    }

    [Theory]
    [InlineData("rows")]
    [InlineData("columns")]
    public async Task Visibility_MapsExplicitWholeDimensionScope(string axis)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeformat", "get-visibility", "--session", "session-1", "--sheet-name", "Sheet1",
            "--range-address", "A1:C3", "--axis", axis
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true,"items":[]}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeformat.get-visibility", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(axis, args.RootElement.GetProperty("axis").GetString(), ignoreCase: true);
        Assert.Equal("A1:C3", args.RootElement.GetProperty("rangeAddress").GetString());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task SetVisibility_MapsExplicitHiddenState(bool hidden)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangeformat", "set-visibility", "--session", "session-1", "--sheet-name", "Sheet1",
            "--range-address", "A1:C3", "--axis", "rows", "--hidden", hidden ? "true" : "false"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("rangeformat.set-visibility", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(hidden, args.RootElement.GetProperty("hidden").GetBoolean());
    }
}
