using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Utilities;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("ProgramTransport")]
[Trait("Layer", "McpServer")]
[Trait("Category", "Integration")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Slow")]
[Trait("RunType", "OnDemand")]
public sealed class SharePointWorkbookTests(ITestOutputHelper output)
    : McpIntegrationTestBase(output, "SharePointWorkbookClient")
{
    [ConfiguredSharePointFact]
    public async Task DirectUrl_EditSaveDiscardAndReopen_PreserveExplicitSaveContract()
    {
        var url = Environment.GetEnvironmentVariable("TEST_SHAREPOINT_WORKBOOK_URL")
            ?? throw new InvalidOperationException("Configured SharePoint fixture is unavailable.");
        var normalizedUrl = FilePathValidation.NormalizeWorkbookLocation(url);
        var sheetName = $"McpTest_{Guid.NewGuid():N}"[..30];
        const string marker = "SharePoint saved marker";
        string? session = null;
        try
        {
            session = await OpenAsync(url);
            var originalSheets = await ReadSheetNamesAsync(session);
            var duplicate = await CallToolAsync("file", new()
            {
                ["action"] = "open",
                ["file_path"] = normalizedUrl,
                ["show"] = true
            });
            using (var rejected = JsonDocument.Parse(duplicate))
            {
                Assert.False(rejected.RootElement.GetProperty("success").GetBoolean());
                Assert.Contains("already open", rejected.RootElement.GetProperty("errorMessage").GetString(), StringComparison.OrdinalIgnoreCase);
            }

            await CreateWorksheetAsync(session, sheetName);
            await WriteMarkerAsync(session, sheetName, "discarded marker");
            await CloseAsync(session, save: false);
            session = null;
            session = await OpenAsync(normalizedUrl);
            Assert.Equal(originalSheets, await ReadSheetNamesAsync(session));

            await CreateWorksheetAsync(session, sheetName);
            await WriteMarkerAsync(session, sheetName, marker);
            await CloseAsync(session, save: true);
            session = null;
            session = await OpenAsync(url);
            var read = await CallToolAsync("range_read", new()
            {
                ["action"] = "get-values",
                ["workbook_session_id"] = session,
                ["sheet_name"] = sheetName,
                ["range_address"] = "A1"
            });
            AssertSuccess(read, "Read saved SharePoint marker");
            using (var values = JsonDocument.Parse(read))
                Assert.Equal(marker, values.RootElement.GetProperty("values")[0][0].GetString());

            var deleted = await CallToolAsync("worksheet", new()
            {
                ["action"] = "delete",
                ["workbook_session_id"] = session,
                ["sheet_name"] = sheetName
            });
            AssertSuccess(deleted, "Remove temporary SharePoint test sheet");
            await CloseAsync(session, save: true);
            session = null;
            session = await OpenAsync(normalizedUrl);
            Assert.Equal(originalSheets, await ReadSheetNamesAsync(session));
        }
        finally
        {
            if (session != null)
                await CloseAsync(session, save: false);
        }
    }

    private async Task<string> OpenAsync(string url)
    {
        var opened = await CallToolAsync("file", new()
        {
            ["action"] = "open",
            ["file_path"] = url,
            ["show"] = true,
            ["timeout_seconds"] = 120
        });
        AssertSuccess(opened, "Open SharePoint workbook");
        using var json = JsonDocument.Parse(opened);
        var session = json.RootElement.GetProperty("workbook_session_id").GetString();
        Assert.False(string.IsNullOrWhiteSpace(session));
        TrackSession(session);
        Assert.Equal(FilePathValidation.NormalizeWorkbookLocation(url), json.RootElement.GetProperty("filePath").GetString());

        var info = await CallToolAsync("workbook_read", new()
        {
            ["action"] = "get-info",
            ["workbook_session_id"] = session
        });
        AssertSuccess(info, "Inspect SharePoint editing rights");
        using var metadata = JsonDocument.Parse(info);
        Assert.False(metadata.RootElement.GetProperty("readOnly").GetBoolean());
        Assert.False(metadata.RootElement.GetProperty("autoSaveOn").GetBoolean());
        return session!;
    }

    private async Task<string[]> ReadSheetNamesAsync(string session)
    {
        var read = await CallToolAsync("worksheet_read", new()
        {
            ["action"] = "list",
            ["workbook_session_id"] = session
        });
        AssertSuccess(read, "List SharePoint worksheets");
        using var json = JsonDocument.Parse(read);
        return json.RootElement.GetProperty("worksheets").EnumerateArray()
            .Select(sheet => sheet.GetProperty("name").GetString()!).ToArray();
    }

    private async Task WriteMarkerAsync(string session, string sheet, string marker)
    {
        var written = await CallToolAsync("range", new()
        {
            ["action"] = "set-values",
            ["workbook_session_id"] = session,
            ["sheet_name"] = sheet,
            ["range_address"] = "A1",
            ["values"] = new string[][] { [marker] }
        });
        AssertSuccess(written, "Write SharePoint test marker");
    }

    private async Task CloseAsync(string session, bool save)
    {
        var closed = await CallToolAsync("file", new()
        {
            ["action"] = "close",
            ["workbook_session_id"] = session,
            ["save"] = save
        });
        AssertSuccess(closed, "Close SharePoint workbook");
        UntrackSession(session);
    }
}

internal sealed class ConfiguredSharePointFactAttribute : FactAttribute
{
    public ConfiguredSharePointFactAttribute()
    {
        if (string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("TEST_SHAREPOINT_WORKBOOK_URL")))
            Skip = "Set TEST_SHAREPOINT_WORKBOOK_URL to a writable SharePoint test workbook. This opt-in test saves a temporary sheet and removes it after verification.";
    }
}
