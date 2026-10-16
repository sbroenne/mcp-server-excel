using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// MCP arguments for simple generated actions whose workbook behavior is covered
/// by the Service tests.
/// </summary>
[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "DataModel")]
[Trait("Feature", "Slicer")]
[Trait("Feature", "Range")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "false")]
public sealed class GeneratedActionMappingProtocolTests(RecordingProgramTransportFixture fixture)
{
    [Fact]
    public async Task DataModelRenameTable_MapsOldAndNewName()
    {
        var call = await fixture.CallToolAsync("datamodel", new()
        {
            ["action"] = "rename-table",
            ["workbook_session_id"] = "session-1",
            ["old_name"] = "ProductData",
            ["new_name"] = "RenamedProductData"
        }, RecordingToolTest.Success("""{"success":true}"""), "datamodel.rename-table",
            """{"oldName":"ProductData","newName":"RenamedProductData"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task SlicerCreate_MapsPivotFieldAndPlacement()
    {
        var call = await fixture.CallToolAsync("slicer", new()
        {
            ["action"] = "create-slicer",
            ["workbook_session_id"] = "session-1",
            ["pivot_table_name"] = "SalesPivot",
            ["field_name"] = "Name",
            ["slicer_name"] = "ReportNames",
            ["destination_sheet"] = "Data",
            ["position"] = "AC1"
        }, RecordingToolTest.Success("""{"success":true}"""), "slicer.create-slicer",
            """{"pivotTableName":"SalesPivot","fieldName":"Name","slicerName":"ReportNames","destinationSheet":"Data","position":"AC1"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task RangeSetStyle_MapsStyleName()
    {
        var call = await fixture.CallToolAsync("range_format", new()
        {
            ["action"] = "set-style",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "Data",
            ["range_address"] = "AD10",
            ["style_name"] = "WorkflowHighlight"
        }, RecordingToolTest.Success("""{"success":true}"""), "rangeformat.set-style",
            """{"sheetName":"Data","rangeAddress":"AD10","styleName":"WorkflowHighlight"}""");
        Assert.False(call.Result.IsError);
    }

    [Fact]
    public async Task WorksheetHide_MapsSheetName()
    {
        var call = await fixture.CallToolAsync("worksheet_style", new()
        {
            ["action"] = "hide",
            ["workbook_session_id"] = "session-1",
            ["sheet_name"] = "InspectorHidden"
        }, RecordingToolTest.Success("""{"success":true}"""), "worksheetstyle.hide",
            """{"sheetName":"InspectorHidden"}""");
        Assert.False(call.Result.IsError);
    }
}
