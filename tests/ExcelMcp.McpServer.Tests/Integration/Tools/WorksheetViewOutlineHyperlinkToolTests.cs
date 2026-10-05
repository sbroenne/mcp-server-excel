using System.Text.Json;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "false")]
public sealed class WorksheetViewOutlineHyperlinkToolTests(
    RecordingProgramTransportFixture fixture)
{
    private const string SessionId = "recording-session";
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task WindowViewActions_RoundTripThroughMcp()
    {
        await AssertSuccessAsync("window", new()
        {
            ["action"] = "freeze-panes",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "View",
            ["frozen_rows"] = 2,
            ["frozen_columns"] = 1
        }, "window.freeze-panes",
        """{"sheetName":"View","frozenRows":2,"frozenColumns":1}""", args =>
        {
            Assert.Equal(2, args.GetProperty("frozenRows").GetInt32());
            Assert.Equal(1, args.GetProperty("frozenColumns").GetInt32());
        });

        var frozenJson = await CallAsync(
            "window",
            new()
            {
                ["action"] = "get-view",
                ["workbook_session_id"] = SessionId,
                ["sheet_name"] = "View"
            },
            "window.get-view",
            """{"sheetName":"View"}""",
            """{"success":true,"freezePanes":true,"splitRow":2,"splitColumn":1}""");
        using (var frozen = JsonDocument.Parse(frozenJson))
        {
            Assert.True(frozen.RootElement.GetProperty("freezePanes").GetBoolean());
            Assert.Equal(2, frozen.RootElement.GetProperty("splitRow").GetInt32());
            Assert.Equal(1, frozen.RootElement.GetProperty("splitColumn").GetInt32());
        }

        await AssertSuccessAsync("window", new()
        {
            ["action"] = "unfreeze-panes",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "View"
        }, "window.unfreeze-panes", """{"sheetName":"View"}""");
        await AssertSuccessAsync("window", new()
        {
            ["action"] = "set-zoom",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "View",
            ["zoom"] = 125
        }, "window.set-zoom", """{"sheetName":"View","zoom":125}""", args =>
            Assert.Equal(125, args.GetProperty("zoom").GetInt32()));
        await AssertSuccessAsync("window", new()
        {
            ["action"] = "set-display-options",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "View",
            ["show_gridlines"] = false,
            ["show_headings"] = false,
            ["show_outline_symbols"] = false,
            ["show_formulas"] = true
        }, "window.set-display-options",
        """{"sheetName":"View","showGridlines":false,"showHeadings":false,"showOutlineSymbols":false,"showFormulas":true}""", args =>
        {
            Assert.False(args.GetProperty("showGridlines").GetBoolean());
            Assert.False(args.GetProperty("showHeadings").GetBoolean());
            Assert.False(args.GetProperty("showOutlineSymbols").GetBoolean());
            Assert.True(args.GetProperty("showFormulas").GetBoolean());
        });
        await AssertSuccessAsync("window", new()
        {
            ["action"] = "set-split",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "View",
            ["split_rows"] = 4,
            ["split_columns"] = 2
        }, "window.set-split",
        """{"sheetName":"View","splitRows":4,"splitColumns":2}""", args =>
        {
            Assert.Equal(4, args.GetProperty("splitRows").GetInt32());
            Assert.Equal(2, args.GetProperty("splitColumns").GetInt32());
        });
    }

    [Fact]
    public async Task WorksheetOutlineActions_RoundTripThroughMcp()
    {
        var missingAxisJson = await _fixture.CallToolWithoutDispatchAsync(
            "worksheet_style",
            new()
            {
                ["action"] = "group",
                ["workbook_session_id"] = SessionId,
                ["sheet_name"] = "Outline",
                ["range_address"] = "2:5"
            });
        using (var result = JsonDocument.Parse(missingAxisJson))
        {
            Assert.False(result.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains(
                "axis",
                result.RootElement.GetProperty("errorMessage").GetString(),
                StringComparison.OrdinalIgnoreCase);
        }

        await AssertSuccessAsync(
            "worksheet_style",
            OutlineArgs("group"),
            "sheet.group",
            """{"sheetName":"Outline","rangeAddress":"2:5","axis":"Rows"}""",
            args => Assert.Equal("Rows", args.GetProperty("axis").GetString()));
        await AssertSuccessAsync(
            "worksheet_style",
            OutlineArgs("ungroup"),
            "sheet.ungroup",
            """{"sheetName":"Outline","rangeAddress":"2:5","axis":"Rows"}""");
        await AssertSuccessAsync("worksheet_style", new()
        {
            ["action"] = "set-outline-settings",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "Outline",
            ["summary_row"] = "above",
            ["summary_column"] = "left",
            ["automatic_styles"] = true
        }, "sheet.set-outline-settings",
        """{"sheetName":"Outline","summaryRow":"above","summaryColumn":"left","automaticStyles":true}""", args =>
        {
            Assert.Equal("above", args.GetProperty("summaryRow").GetString());
            Assert.Equal("left", args.GetProperty("summaryColumn").GetString());
            Assert.True(args.GetProperty("automaticStyles").GetBoolean());
        });
        await AssertSuccessAsync("worksheet_style", new()
        {
            ["action"] = "show-outline-levels",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "Outline",
            ["row_levels"] = 1
        }, "sheet.show-outline-levels",
            """{"sheetName":"Outline","rowLevels":1}""",
            args => Assert.Equal(1, args.GetProperty("rowLevels").GetInt32()));

        var infoJson = await CallAsync(
            "worksheet_style_read",
            OutlineArgs("get-outline-info"),
            "sheet.get-outline-info",
            """{"sheetName":"Outline","rangeAddress":"2:5","axis":"Rows"}""",
            """{"success":true,"outlineLevel":2,"summaryRow":"above","summaryColumn":"left","hidden":true}""");
        using (var info = JsonDocument.Parse(infoJson))
        {
            Assert.Equal(2, info.RootElement.GetProperty("outlineLevel").GetInt32());
            Assert.Equal("above", info.RootElement.GetProperty("summaryRow").GetString());
            Assert.Equal("left", info.RootElement.GetProperty("summaryColumn").GetString());
            Assert.True(info.RootElement.GetProperty("hidden").GetBoolean());
        }

        await AssertSuccessAsync("worksheet_style", new()
        {
            ["action"] = "clear-outline",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "Outline"
        }, "sheet.clear-outline", """{"sheetName":"Outline"}""");
    }

    [Fact]
    public async Task InternalHyperlinkLifecycle_RoundTripsThroughMcp()
    {
        await AssertSuccessAsync("range_link", new()
        {
            ["action"] = "add-hyperlink",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "Links",
            ["cell_address"] = "A1",
            ["sub_address"] = "'Links'!D5",
            ["display_text"] = "Jump"
        }, "rangelink.add-hyperlink",
        """{"sheetName":"Links","cellAddress":"A1","subAddress":"'Links'!D5","displayText":"Jump"}""", args =>
        {
            Assert.Equal("'Links'!D5", args.GetProperty("subAddress").GetString());
            Assert.Equal("Jump", args.GetProperty("displayText").GetString());
        });

        var listJson = await CallAsync(
            "range_link_read",
            new()
            {
                ["action"] = "list-hyperlinks",
                ["workbook_session_id"] = SessionId,
                ["sheet_name"] = "Links"
            },
            "rangelink.list-hyperlinks",
            """{"sheetName":"Links"}""",
            """{"success":true,"hyperlinks":[{"isInternal":true,"subAddress":"'Links'!D5"}]}""");
        using (var list = JsonDocument.Parse(listJson))
        {
            var hyperlink = list.RootElement.GetProperty("hyperlinks")[0];
            Assert.True(hyperlink.GetProperty("isInternal").GetBoolean());
            Assert.Equal("'Links'!D5", hyperlink.GetProperty("subAddress").GetString());
        }

        await AssertSuccessAsync("range_link", new()
        {
            ["action"] = "update-hyperlink",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "Links",
            ["cell_address"] = "A1",
            ["url"] = "https://example.com",
            ["sub_address"] = "target",
            ["display_text"] = "Updated",
            ["tooltip"] = "Updated link"
        }, "rangelink.update-hyperlink",
        """{"sheetName":"Links","cellAddress":"A1","url":"https://example.com","displayText":"Updated","tooltip":"Updated link","subAddress":"target"}""", args =>
        {
            Assert.Equal("https://example.com", args.GetProperty("url").GetString());
            Assert.Equal("target", args.GetProperty("subAddress").GetString());
            Assert.Equal("Updated", args.GetProperty("displayText").GetString());
            Assert.Equal("Updated link", args.GetProperty("tooltip").GetString());
        });

        await AssertSuccessAsync("range_link", new()
        {
            ["action"] = "remove-hyperlink",
            ["workbook_session_id"] = SessionId,
            ["sheet_name"] = "Links",
            ["range_address"] = "A1"
        }, "rangelink.remove-hyperlink",
            """{"sheetName":"Links","rangeAddress":"A1"}""",
            args => Assert.Equal("A1", args.GetProperty("rangeAddress").GetString()));
    }

    private static Dictionary<string, object?> OutlineArgs(string action) => new()
    {
        ["action"] = action,
        ["workbook_session_id"] = SessionId,
        ["sheet_name"] = "Outline",
        ["range_address"] = "2:5",
        ["axis"] = "Rows"
    };

    private async Task AssertSuccessAsync(
        string tool,
        Dictionary<string, object?> arguments,
        string command,
        string? expectedArgsJson,
        Action<JsonElement>? assertArgs = null)
    {
        var json = await CallAsync(
            tool,
            arguments,
            command,
            expectedArgsJson,
            """{"success":true}""",
            assertArgs);
        using var result = JsonDocument.Parse(json);
        Assert.True(result.RootElement.GetProperty("success").GetBoolean());
    }

    private async Task<string> CallAsync(
        string tool,
        Dictionary<string, object?> arguments,
        string command,
        string? expectedArgsJson,
        string responseJson,
        Action<JsonElement>? assertArgs = null)
    {
        var call = await _fixture.CallToolAsync(
            tool,
            arguments,
            RecordingToolTest.Success(responseJson),
            command,
            expectedArgsJson);
        using var args = RecordingToolTest.ParseArgs(
            call.Request,
            command,
            SessionId);
        Assert.Equal(
            arguments["sheet_name"],
            args.RootElement.GetProperty("sheetName").GetString());
        assertArgs?.Invoke(args.RootElement);
        return call.JsonResult;
    }
}
