using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Feature", "Protection")]
[Trait("Layer", "McpServer")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class FineProtectionProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task SheetProtection_PreservesNestedNativePermissionNames()
    {
        var options = new { allowFiltering = true, allowFormattingRows = true, userInterfaceOnly = true };
        var expected = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            isProtected = true,
            options = new SheetProtectionOptions
            {
                AllowFiltering = true,
                AllowFormattingRows = true,
                UserInterfaceOnly = true
            }
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("worksheet_style", new()
        {
            ["action"] = "set-protection",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["is_protected"] = true,
            ["options"] = options
        }, RecordingToolTest.Success("""{"success":true}"""), "sheet.set-protection", expected);
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public async Task CellProtection_MapsExplicitFlagsAndExactScope(bool locked, bool hidden)
    {
        var expected = JsonSerializer.Serialize(new
        {
            sheetName = "Sheet1",
            rangeAddress = "A1,A3",
            locked,
            formulaHidden = hidden
        }, ServiceProtocol.JsonOptions);
        var call = await fixture.CallToolAsync("range_link", new()
        {
            ["action"] = "set-cell-protection",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1,A3",
            ["locked"] = locked,
            ["formula_hidden"] = hidden
        }, RecordingToolTest.Success("""{"success":true}"""), "rangelink.set-cell-protection", expected);
        Assert.False(call.Result.IsError);
    }

    [Theory]
    [InlineData("get-cell-lock")]
    [InlineData("set-cell-lock")]
    public async Task RemovedLockActions_DoNotDispatch(string action)
    {
        var result = await fixture.CallResultWithoutDispatchAsync("range_link", new()
        {
            ["action"] = action,
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Sheet1",
            ["range_address"] = "A1"
        });
        Assert.True(result.IsError);
    }

    [Fact]
    public async Task Discovery_DescribesRuntimeOnlyProtectionAndCompleteCellReads()
    {
        var tools = await fixture.ListToolsAsync();
        var sheet = Assert.Single(tools, item => item.Name == "worksheet_style");
        Assert.Contains("runtime-only", sheet.Description, StringComparison.Ordinal);
        Assert.True(sheet.JsonSchema.GetProperty("properties").TryGetProperty("options", out _));
        var range = Assert.Single(tools, item => item.Name == "range_link");
        Assert.Contains("not tool inspection or file encryption", range.Description, StringComparison.Ordinal);
        var readRange = Assert.Single(tools, item => item.Name == "range_link_read");
        Assert.Contains("every unique requested cell, without a cap", readRange.Description, StringComparison.Ordinal);
    }
}
