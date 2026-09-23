using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Exercises worksheet view, outline, and hyperlink operations through the real MCP protocol.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "Worksheets")]
public sealed class WorksheetViewOutlineHyperlinkToolTests : McpIntegrationTestBase
{
    private readonly string _tempDirectory;

    public WorksheetViewOutlineHyperlinkToolTests(ITestOutputHelper output)
        : base(output, "WorksheetViewOutlineHyperlinkClient")
    {
        _tempDirectory = CreateTempDirectory("WorksheetViewOutlineHyperlink");
    }

    [Fact]
    public async Task WindowViewActions_RoundTripThroughMcp()
    {
        var sessionId = await CreateWorkbookSessionAsync(Path.Join(_tempDirectory, "WindowView.xlsx"));
        await CreateWorksheetAsync(sessionId, "View");

        AssertSuccess(await CallToolAsync("window", new()
        {
            ["action"] = "freeze-panes",
            ["session_id"] = sessionId,
            ["sheet_name"] = "View",
            ["frozen_rows"] = 2,
            ["frozen_columns"] = 1
        }), "window.freeze-panes");

        using (var frozen = ParseJsonResult(await GetViewAsync(sessionId), "window.get-view"))
        {
            Assert.True(frozen.RootElement.GetProperty("freezePanes").GetBoolean());
            Assert.Equal(2, frozen.RootElement.GetProperty("splitRow").GetInt32());
            Assert.Equal(1, frozen.RootElement.GetProperty("splitColumn").GetInt32());
        }

        AssertSuccess(await CallToolAsync("window", new()
        {
            ["action"] = "unfreeze-panes",
            ["session_id"] = sessionId,
            ["sheet_name"] = "View"
        }), "window.unfreeze-panes");
        using (var unfrozen = ParseJsonResult(await GetViewAsync(sessionId), "window.get-view unfrozen"))
        {
            Assert.False(unfrozen.RootElement.GetProperty("freezePanes").GetBoolean());
            Assert.Equal(0, unfrozen.RootElement.GetProperty("splitRow").GetInt32());
            Assert.Equal(0, unfrozen.RootElement.GetProperty("splitColumn").GetInt32());
        }

        AssertSuccess(await CallToolAsync("window", new()
        {
            ["action"] = "set-zoom",
            ["session_id"] = sessionId,
            ["sheet_name"] = "View",
            ["zoom"] = 125
        }), "window.set-zoom");
        AssertSuccess(await CallToolAsync("window", new()
        {
            ["action"] = "set-display-options",
            ["session_id"] = sessionId,
            ["sheet_name"] = "View",
            ["show_gridlines"] = false,
            ["show_headings"] = false,
            ["show_outline_symbols"] = false,
            ["show_formulas"] = true
        }), "window.set-display-options");
        AssertSuccess(await CallToolAsync("window", new()
        {
            ["action"] = "freeze-panes",
            ["session_id"] = sessionId,
            ["sheet_name"] = "View",
            ["frozen_rows"] = 2,
            ["frozen_columns"] = 1
        }), "window.freeze-panes before split");
        AssertSuccess(await CallToolAsync("window", new()
        {
            ["action"] = "set-split",
            ["session_id"] = sessionId,
            ["sheet_name"] = "View",
            ["split_rows"] = 4,
            ["split_columns"] = 2
        }), "window.set-split");
        using (var split = ParseJsonResult(await GetViewAsync(sessionId), "window.get-view split"))
        {
            Assert.Equal(4, split.RootElement.GetProperty("splitRow").GetInt32());
            Assert.Equal(2, split.RootElement.GetProperty("splitColumn").GetInt32());
        }

        using var view = ParseJsonResult(await GetViewAsync(sessionId), "window.get-view final");
        var root = view.RootElement;
        Assert.False(root.GetProperty("freezePanes").GetBoolean());
        Assert.Equal(4, root.GetProperty("splitRow").GetInt32());
        Assert.Equal(2, root.GetProperty("splitColumn").GetInt32());
        Assert.Equal(125, root.GetProperty("zoom").GetInt32());
        Assert.False(root.GetProperty("displayGridlines").GetBoolean());
        Assert.False(root.GetProperty("displayHeadings").GetBoolean());
        Assert.False(root.GetProperty("displayOutlineSymbols").GetBoolean());
        Assert.True(root.GetProperty("displayFormulas").GetBoolean());
    }

    [Fact]
    public async Task WorksheetOutlineActions_RoundTripThroughMcp()
    {
        var sessionId = await CreateWorkbookSessionAsync(Path.Join(_tempDirectory, "Outline.xlsx"));
        await CreateWorksheetAsync(sessionId, "Outline");

        var missingAxisJson = await CallToolAsync("worksheet_style", new()
        {
            ["action"] = "group",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Outline",
            ["range_address"] = "2:5"
        });
        using (var missingAxis = ParseJsonResult(missingAxisJson, "worksheet_style.group missing axis"))
        {
            Assert.False(missingAxis.RootElement.GetProperty("success").GetBoolean());
            Assert.Contains("axis", missingAxisJson, StringComparison.OrdinalIgnoreCase);
        }

        AssertSuccess(await CallToolAsync("worksheet_style", new()
        {
            ["action"] = "group",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Outline",
            ["range_address"] = "2:5",
            ["axis"] = "Rows"
        }), "worksheet_style.group");
        AssertSuccess(await CallToolAsync("worksheet_style", new()
        {
            ["action"] = "ungroup",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Outline",
            ["range_address"] = "2:5",
            ["axis"] = "Rows"
        }), "worksheet_style.ungroup");
        AssertSuccess(await CallToolAsync("worksheet_style", new()
        {
            ["action"] = "group",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Outline",
            ["range_address"] = "2:5",
            ["axis"] = "Rows"
        }), "worksheet_style.group after ungroup");
        AssertSuccess(await CallToolAsync("worksheet_style", new()
        {
            ["action"] = "set-outline-settings",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Outline",
            ["summary_row"] = "above",
            ["summary_column"] = "left",
            ["automatic_styles"] = true
        }), "worksheet_style.set-outline-settings");
        AssertSuccess(await CallToolAsync("worksheet_style", new()
        {
            ["action"] = "show-outline-levels",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Outline",
            ["row_levels"] = 1
        }), "worksheet_style.show-outline-levels");

        var infoJson = await CallToolAsync("worksheet_style", new()
        {
            ["action"] = "get-outline-info",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Outline",
            ["range_address"] = "2:5",
            ["axis"] = "Rows"
        });
        using var info = ParseJsonResult(infoJson, "worksheet_style.get-outline-info");
        Assert.Equal(2, info.RootElement.GetProperty("outlineLevel").GetInt32());
        Assert.Equal("above", info.RootElement.GetProperty("summaryRow").GetString());
        Assert.Equal("left", info.RootElement.GetProperty("summaryColumn").GetString());
        Assert.True(info.RootElement.GetProperty("hidden").GetBoolean());

        AssertSuccess(await CallToolAsync("worksheet_style", new()
        {
            ["action"] = "clear-outline",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Outline"
        }), "worksheet_style.clear-outline");
    }

    [Fact]
    public async Task InternalHyperlinkLifecycle_RoundTripsThroughMcp()
    {
        var sessionId = await CreateWorkbookSessionAsync(Path.Join(_tempDirectory, "Hyperlink.xlsx"));
        await CreateWorksheetAsync(sessionId, "Links");

        AssertSuccess(await CallToolAsync("range_link", new()
        {
            ["action"] = "add-hyperlink",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Links",
            ["cell_address"] = "A1",
            ["sub_address"] = "'Links'!D5",
            ["display_text"] = "Jump"
        }), "range_link.add-hyperlink");
        var listJson = await CallToolAsync("range_link", new()
        {
            ["action"] = "list-hyperlinks",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Links"
        });
        using (var list = ParseJsonResult(listJson, "range_link.list-hyperlinks"))
        {
            var hyperlink = list.RootElement.GetProperty("hyperlinks")[0];
            Assert.True(hyperlink.GetProperty("isInternal").GetBoolean());
            Assert.Equal("'Links'!D5", hyperlink.GetProperty("subAddress").GetString());
        }
        AssertSuccess(await CallToolAsync("range_link", new()
        {
            ["action"] = "update-hyperlink",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Links",
            ["cell_address"] = "A1",
            ["url"] = "https://example.com",
            ["sub_address"] = "target",
            ["display_text"] = "Updated",
            ["tooltip"] = "Updated link"
        }), "range_link.update-hyperlink");

        var getJson = await CallToolAsync("range_link", new()
        {
            ["action"] = "get-hyperlink",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Links",
            ["cell_address"] = "A1"
        });
        using (var get = ParseJsonResult(getJson, "range_link.get-hyperlink"))
        {
            var hyperlink = get.RootElement.GetProperty("hyperlinks")[0];
            Assert.StartsWith("https://example.com", hyperlink.GetProperty("address").GetString());
            Assert.Equal("target", hyperlink.GetProperty("subAddress").GetString());
            Assert.Equal("Updated", hyperlink.GetProperty("displayText").GetString());
        }

        AssertSuccess(await CallToolAsync("range_link", new()
        {
            ["action"] = "remove-hyperlink",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Links",
            ["range_address"] = "A1"
        }), "range_link.remove-hyperlink");
    }

    private Task<string> GetViewAsync(string sessionId)
    {
        return CallToolAsync("window", new()
        {
            ["action"] = "get-view",
            ["session_id"] = sessionId,
            ["sheet_name"] = "View"
        });
    }
}
