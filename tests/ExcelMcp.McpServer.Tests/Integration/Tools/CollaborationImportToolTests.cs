using System.Text.Json;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

/// <summary>
/// Black-box MCP coverage for local collaboration, QueryTable import, and refresh-control actions.
/// </summary>
[Collection("ProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "CollaborationImport")]
[Trait("RequiresExcel", "true")]
public sealed class CollaborationImportToolTests(ITestOutputHelper output)
    : McpIntegrationTestBase(output, "CollaborationImportClient")
{
    private static readonly TimeSpan ToolTimeout = TimeSpan.FromSeconds(90);

    [Fact]
    public async Task ThreadedComments_AllActionsMetadataValidationAndCleanup_ViaMcp()
    {
        var tempDirectory = CreateTempDirectory("McpThreadedComments");
        var workbookPath = Path.Join(tempDirectory, "ThreadedComments.xlsx");
        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await CreateWorksheetAsync(sessionId, "Review");

        var addJson = await CallToolAsync("range_link", new Dictionary<string, object?>
        {
            ["action"] = "add-threaded-comment",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Review",
            ["cell_address"] = "B2",
            ["text"] = "Review this value"
        }, ToolTimeout);
        AssertSuccess(addJson, "range_link.add-threaded-comment");

        var duplicateJson = await CallToolAsync("range_link", new Dictionary<string, object?>
        {
            ["action"] = "add-threaded-comment",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Review",
            ["cell_address"] = "B2",
            ["text"] = "Duplicate"
        }, ToolTimeout);
        AssertFailure(duplicateJson, "range_link.add-threaded-comment duplicate");

        var replyJson = await CallToolAsync("range_link", new Dictionary<string, object?>
        {
            ["action"] = "add-threaded-comment-reply",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Review",
            ["cell_address"] = "B2",
            ["text"] = "Reviewed"
        }, ToolTimeout);
        AssertSuccess(replyJson, "range_link.add-threaded-comment-reply");

        var listJson = await CallToolAsync("range_link", new Dictionary<string, object?>
        {
            ["action"] = "list-threaded-comments",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Review",
            ["cell_address"] = "B2"
        }, ToolTimeout);
        AssertSuccess(listJson, "range_link.list-threaded-comments");
        using (var listDocument = JsonDocument.Parse(listJson))
        {
            var comment = Assert.Single(listDocument.RootElement.GetProperty("comments").EnumerateArray());
            Assert.Equal("B2", comment.GetProperty("cellAddress").GetString());
            Assert.Equal("Review this value", comment.GetProperty("text").GetString());
            Assert.False(string.IsNullOrWhiteSpace(comment.GetProperty("authorName").GetString()));
            var reply = Assert.Single(comment.GetProperty("replies").EnumerateArray());
            Assert.Equal("Reviewed", reply.GetProperty("text").GetString());
            Assert.False(string.IsNullOrWhiteSpace(reply.GetProperty("authorName").GetString()));
        }

        var deleteJson = await CallToolAsync("range_link", new Dictionary<string, object?>
        {
            ["action"] = "delete-threaded-comment",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Review",
            ["cell_address"] = "B2"
        }, ToolTimeout);
        AssertSuccess(deleteJson, "range_link.delete-threaded-comment");

        var finalListJson = await CallToolAsync("range_link", new Dictionary<string, object?>
        {
            ["action"] = "list-threaded-comments",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Review",
            ["cell_address"] = "B2"
        }, ToolTimeout);
        AssertSuccess(finalListJson, "range_link.list-threaded-comments after delete");
        using var finalListDocument = JsonDocument.Parse(finalListJson);
        Assert.Empty(finalListDocument.RootElement.GetProperty("comments").EnumerateArray());
    }

    [Fact]
    public async Task QueryTable_AllActionsLifecycleMetadataValidationAndCleanup_ViaMcp()
    {
        var tempDirectory = CreateTempDirectory("McpQueryTable");
        var workbookPath = Path.Join(tempDirectory, "QueryTables.xlsx");
        var csvPath = Path.Join(
            tempDirectory,
            "orders;User ID=review-user-id;Password={review-secret;credential-tail};UID=review-uid;PWD=review-pwd.csv");
        var htmlPath = Path.Join(tempDirectory, "rates.html");
        await File.WriteAllTextAsync(csvPath, "Name,Value\nCafé,10\nBeta,20\n");
        await File.WriteAllTextAsync(
            htmlPath,
            "<html><body><table><tr><th>Name</th><th>Value</th></tr><tr><td>Alpha</td><td>10</td></tr></table></body></html>");

        var sessionId = await CreateWorkbookSessionAsync(workbookPath);
        await CreateWorksheetAsync(sessionId, "Imports");

        var invalidCreateJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "create-text",
            ["session_id"] = sessionId,
            ["query_table_name"] = "InvalidImport",
            ["source_path"] = csvPath,
            ["sheet_name"] = "Imports",
            ["destination_address"] = "A1",
            ["delimiter"] = ",,"
        }, ToolTimeout);
        AssertFailure(invalidCreateJson, "querytable.create-text invalid delimiter");

        var createTextJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "create-text",
            ["session_id"] = sessionId,
            ["query_table_name"] = "CsvImport",
            ["source_path"] = csvPath,
            ["sheet_name"] = "Imports",
            ["destination_address"] = "B2",
            ["delimiter"] = ",",
            ["text_qualifier"] = "double-quote",
            ["encoding"] = 65001,
            ["has_headers"] = true
        }, ToolTimeout);
        AssertSuccess(createTextJson, "querytable.create-text");

        var listJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, ToolTimeout);
        AssertSuccess(listJson, "querytable.list");
        using (var listDocument = JsonDocument.Parse(listJson))
        {
            var item = Assert.Single(listDocument.RootElement.GetProperty("queryTables").EnumerateArray());
            Assert.Equal("CsvImport", item.GetProperty("name").GetString());
            Assert.Equal("Imports", item.GetProperty("sheetName").GetString());
            Assert.Equal("B2", item.GetProperty("destination").GetString());
            Assert.Equal("text", item.GetProperty("sourceType").GetString());
        }

        var viewJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "view",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "CsvImport"
        }, ToolTimeout);
        AssertSuccess(viewJson, "querytable.view text");
        using (var viewDocument = JsonDocument.Parse(viewJson))
        {
            var root = viewDocument.RootElement;
            Assert.Equal(",", root.GetProperty("delimiter").GetString());
            Assert.True(root.GetProperty("encoding").GetInt32() > 0);
            Assert.Equal("text", root.GetProperty("sourceType").GetString());
            var connection = root.GetProperty("connection").GetString();
            Assert.DoesNotContain("review-user-id", connection, StringComparison.Ordinal);
            Assert.DoesNotContain("review-secret", connection, StringComparison.Ordinal);
            Assert.DoesNotContain("credential-tail", connection, StringComparison.Ordinal);
            Assert.DoesNotContain("review-uid", connection, StringComparison.Ordinal);
            Assert.DoesNotContain("review-pwd", connection, StringComparison.Ordinal);
            Assert.Contains("(redacted)", connection, StringComparison.Ordinal);
        }

        var importedValuesJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "get-values",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["range_address"] = "B3"
        }, ToolTimeout);
        AssertSuccess(importedValuesJson, "range.get-values imported UTF-8 text");
        using (var importedValuesDocument = JsonDocument.Parse(importedValuesJson))
        {
            Assert.Equal(
                "Café",
                importedValuesDocument.RootElement.GetProperty("values")[0][0].GetString());
        }

        var invalidPropertiesJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "set-properties",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "CsvImport",
            ["refresh_period"] = -1
        }, ToolTimeout);
        AssertFailure(invalidPropertiesJson, "querytable.set-properties invalid refresh period");

        var setPropertiesJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "set-properties",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "CsvImport",
            ["background_query"] = false,
            ["refresh_on_file_open"] = true,
            ["refresh_period"] = 15,
            ["adjust_column_width"] = false,
            ["preserve_formatting"] = true
        }, ToolTimeout);
        AssertSuccess(setPropertiesJson, "querytable.set-properties");

        var refreshJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "refresh",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "CsvImport"
        }, ToolTimeout);
        AssertSuccess(refreshJson, "querytable.refresh");

        var statusJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "get-refresh-status",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "CsvImport"
        }, ToolTimeout);
        AssertSuccess(statusJson, "querytable.get-refresh-status");
        using (var statusDocument = JsonDocument.Parse(statusJson))
        {
            Assert.True(statusDocument.RootElement.GetProperty("supportsRefreshStatus").GetBoolean());
            Assert.False(statusDocument.RootElement.GetProperty("isRefreshing").GetBoolean());
        }

        var cancelJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "cancel-refresh",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "CsvImport"
        }, ToolTimeout);
        AssertSuccess(cancelJson, "querytable.cancel-refresh");
        using (var cancelDocument = JsonDocument.Parse(cancelJson))
        {
            Assert.True(cancelDocument.RootElement.GetProperty("supportsCancellation").GetBoolean());
            Assert.False(cancelDocument.RootElement.GetProperty("wasRefreshing").GetBoolean());
            Assert.False(cancelDocument.RootElement.GetProperty("cancelled").GetBoolean());
        }

        var deleteTextJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "delete",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "CsvImport"
        }, ToolTimeout);
        AssertSuccess(deleteTextJson, "querytable.delete text");

        var createWebJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "create-web",
            ["session_id"] = sessionId,
            ["query_table_name"] = "HtmlImport",
            ["url"] = new Uri(htmlPath).AbsoluteUri,
            ["sheet_name"] = "Imports",
            ["destination_address"] = "A1",
            ["selection_type"] = "specified-tables",
            ["web_tables"] = "1",
            ["formatting"] = "none"
        }, ToolTimeout);
        AssertSuccess(createWebJson, "querytable.create-web");

        var viewWebJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "view",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "HtmlImport"
        }, ToolTimeout);
        AssertSuccess(viewWebJson, "querytable.view web");
        using (var viewWebDocument = JsonDocument.Parse(viewWebJson))
        {
            var root = viewWebDocument.RootElement;
            Assert.Equal("web", root.GetProperty("sourceType").GetString());
            Assert.Equal("specified-tables", root.GetProperty("webSelectionType").GetString());
            Assert.Equal("none", root.GetProperty("webFormatting").GetString());
            Assert.Contains("1", root.GetProperty("webTables").GetString(), StringComparison.Ordinal);
        }

        var deleteWebJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "delete",
            ["session_id"] = sessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "HtmlImport"
        }, ToolTimeout);
        AssertSuccess(deleteWebJson, "querytable.delete web");

        var finalListJson = await CallToolAsync("querytable", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sessionId
        }, ToolTimeout);
        AssertSuccess(finalListJson, "querytable.list after cleanup");
        using var finalListDocument = JsonDocument.Parse(finalListJson);
        Assert.Empty(finalListDocument.RootElement.GetProperty("queryTables").EnumerateArray());
    }

    [Fact]
    public async Task ConnectionRefreshControl_AllActionsMetadataValidationAndCleanup_ViaMcp()
    {
        var tempDirectory = CreateTempDirectory("McpConnectionRefresh");
        var sourceWorkbookPath = Path.Join(tempDirectory, "Source.xlsx");
        var targetWorkbookPath = Path.Join(tempDirectory, "Target.xlsx");

        var sourceSessionId = await CreateWorkbookSessionAsync(sourceWorkbookPath);
        var sourceSheetListJson = await CallToolAsync("worksheet", new Dictionary<string, object?>
        {
            ["action"] = "list",
            ["session_id"] = sourceSessionId
        }, ToolTimeout);
        AssertSuccess(sourceSheetListJson, "worksheet.list source");
        using var sourceSheetDocument = JsonDocument.Parse(sourceSheetListJson);
        var sourceSheetName = sourceSheetDocument.RootElement
            .GetProperty("worksheets")[0]
            .GetProperty("name")
            .GetString();
        Assert.False(string.IsNullOrWhiteSpace(sourceSheetName));

        var writeSourceJson = await CallToolAsync("range", new Dictionary<string, object?>
        {
            ["action"] = "set-values",
            ["session_id"] = sourceSessionId,
            ["sheet_name"] = sourceSheetName,
            ["range_address"] = "A1:B3",
            ["values"] = new object?[][]
            {
                ["Product", "Price"],
                ["Widget", 19.99],
                ["Gadget", 29.99]
            }
        }, ToolTimeout);
        AssertSuccess(writeSourceJson, "range.set-values source");
        await CloseSessionAsync(sourceSessionId, save: true);

        var targetSessionId = await CreateWorkbookSessionAsync(targetWorkbookPath);
        const string connectionName = "ProductsConnection";
        var connectionString =
            $"OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;Data Source={sourceWorkbookPath};Extended Properties=\"Excel 12.0 Xml;HDR=YES\"";
        var createConnectionJson = await CallToolAsync("connection", new Dictionary<string, object?>
        {
            ["action"] = "create",
            ["session_id"] = targetSessionId,
            ["connection_name"] = connectionName,
            ["connection_string"] = connectionString,
            ["command_text"] = $"SELECT * FROM [{sourceSheetName}$]"
        }, ToolTimeout);
        AssertSuccess(createConnectionJson, "connection.create");

        var statusJson = await CallToolAsync("connection", new Dictionary<string, object?>
        {
            ["action"] = "get-refresh-status",
            ["session_id"] = targetSessionId,
            ["connection_name"] = connectionName
        }, ToolTimeout);
        AssertSuccess(statusJson, "connection.get-refresh-status");
        using (var statusDocument = JsonDocument.Parse(statusJson))
        {
            Assert.True(statusDocument.RootElement.GetProperty("supportsRefreshStatus").GetBoolean());
            Assert.False(statusDocument.RootElement.GetProperty("isRefreshing").GetBoolean());
        }

        var cancelJson = await CallToolAsync("connection", new Dictionary<string, object?>
        {
            ["action"] = "cancel-refresh",
            ["session_id"] = targetSessionId,
            ["connection_name"] = connectionName
        }, ToolTimeout);
        AssertSuccess(cancelJson, "connection.cancel-refresh");
        using (var cancelDocument = JsonDocument.Parse(cancelJson))
        {
            Assert.True(cancelDocument.RootElement.GetProperty("supportsCancellation").GetBoolean());
            Assert.False(cancelDocument.RootElement.GetProperty("wasRefreshing").GetBoolean());
            Assert.False(cancelDocument.RootElement.GetProperty("cancelled").GetBoolean());
        }

        var missingStatusJson = await CallToolAsync("connection", new Dictionary<string, object?>
        {
            ["action"] = "get-refresh-status",
            ["session_id"] = targetSessionId,
            ["connection_name"] = "MissingConnection"
        }, ToolTimeout);
        AssertFailure(missingStatusJson, "connection.get-refresh-status missing connection");

        var deleteConnectionJson = await CallToolAsync("connection", new Dictionary<string, object?>
        {
            ["action"] = "delete",
            ["session_id"] = targetSessionId,
            ["connection_name"] = connectionName
        }, ToolTimeout);
        AssertSuccess(deleteConnectionJson, "connection.delete cleanup");
    }

    private static void AssertFailure(string json, string operation)
    {
        using var document = JsonDocument.Parse(json);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean(), $"{operation} unexpectedly succeeded: {json}");
        Assert.False(
            string.IsNullOrWhiteSpace(document.RootElement.GetProperty("errorMessage").GetString()),
            $"{operation} should return an actionable error: {json}");
    }
}
