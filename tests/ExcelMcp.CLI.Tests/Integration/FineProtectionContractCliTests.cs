using System.Text.Json;
using Sbroenne.ExcelMcp.CLI.Tests.Helpers;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

[Collection("Service")]
[Trait("Category", "Integration")]
[Trait("Feature", "Protection")]
[Trait("Layer", "CLI")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class FineProtectionContractCliTests
{
    [Fact]
    public async Task SheetProtection_MapsNativeOptionsJson()
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "worksheetstyle", "set-protection", "--session", "session-1", "--sheet-name", "Sheet1",
            "--is-protected", "true", "--options", """{"allowFiltering":true,"userInterfaceOnly":true}"""
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        Assert.Equal("sheet.set-protection", captured.Command);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.True(args.RootElement.GetProperty("options").GetProperty("allowFiltering").GetBoolean());
        Assert.True(args.RootElement.GetProperty("options").GetProperty("userInterfaceOnly").GetBoolean());
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public async Task CellProtection_MapsExplicitFlags(bool locked, bool hidden)
    {
        ServiceRequest? captured = null;
        var result = await InProcessCliHelper.RunAsync(
        [
            "rangelink", "set-cell-protection", "--session", "session-1",
            "--sheet-name", "Sheet1", "--range-address", "A1,A3",
            "--locked", locked ? "true" : "false", "--formula-hidden", hidden ? "true" : "false"
        ], request =>
        {
            captured = request;
            return new ServiceResponse { Success = true, Result = """{"success":true}""" };
        });
        Assert.True(result.ExitCode == 0, result.Stdout + result.Stderr);
        Assert.NotNull(captured);
        using var args = JsonDocument.Parse(captured.Args!);
        Assert.Equal(locked, args.RootElement.GetProperty("locked").GetBoolean());
        Assert.Equal(hidden, args.RootElement.GetProperty("formulaHidden").GetBoolean());
    }
}
