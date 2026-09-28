using System.Text.Json;
using Sbroenne.ExcelMcp.Service;
using Xunit;

namespace Sbroenne.ExcelMcp.McpServer.Tests.Integration.Tools;

[Collection("RecordingProgramTransport")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "McpServer")]
[Trait("Feature", "CollaborationImport")]
[Trait("RequiresExcel", "false")]
public sealed class CollaborationImportToolTests(
    RecordingProgramTransportFixture fixture)
{
    private const string SessionId = "recording-session";
    private readonly RecordingProgramTransportFixture _fixture = fixture;

    [Fact]
    public async Task ThreadedComments_AllActionsMetadataValidationAndCleanup_ViaMcp()
    {
        await AssertSuccessAsync("range_link", CommentArgs(
            "add-threaded-comment",
            "Review this value"), "rangelink.add-threaded-comment",
            """{"sheetName":"Review","cellAddress":"B2","text":"Review this value"}""");

        var duplicate = await _fixture.CallToolAsync(
            "range_link",
            CommentArgs("add-threaded-comment", "Duplicate"),
            Failure(
                "rangelink.add-threaded-comment",
                "A threaded comment already exists at B2."),
            "rangelink.add-threaded-comment",
            """{"sheetName":"Review","cellAddress":"B2","text":"Duplicate"}""");
        AssertRequest(duplicate.Request, "rangelink.add-threaded-comment");
        AssertFailure(duplicate.JsonResult);

        await AssertSuccessAsync("range_link", CommentArgs(
            "add-threaded-comment-reply",
            "Reviewed"), "rangelink.add-threaded-comment-reply",
            """{"sheetName":"Review","cellAddress":"B2","text":"Reviewed"}""");

        var listJson = await CallAsync(
            "range_link",
            new()
            {
                ["action"] = "list-threaded-comments",
                ["session_id"] = SessionId,
                ["sheet_name"] = "Review",
                ["cell_address"] = "B2"
            },
            "rangelink.list-threaded-comments",
            """{"sheetName":"Review","cellAddress":"B2"}""",
            """
            {
              "success": true,
              "comments": [{
                "cellAddress": "B2",
                "text": "Review this value",
                "authorName": "Test Author",
                "replies": [{"text":"Reviewed","authorName":"Test Author"}]
              }]
            }
            """);
        using (var list = JsonDocument.Parse(listJson))
        {
            var comment = Assert.Single(
                list.RootElement.GetProperty("comments").EnumerateArray());
            Assert.Equal("B2", comment.GetProperty("cellAddress").GetString());
            Assert.Equal(
                "Review this value",
                comment.GetProperty("text").GetString());
            Assert.False(string.IsNullOrWhiteSpace(
                comment.GetProperty("authorName").GetString()));
            var reply = Assert.Single(
                comment.GetProperty("replies").EnumerateArray());
            Assert.Equal("Reviewed", reply.GetProperty("text").GetString());
        }

        await AssertSuccessAsync("range_link", new()
        {
            ["action"] = "delete-threaded-comment",
            ["session_id"] = SessionId,
            ["sheet_name"] = "Review",
            ["cell_address"] = "B2"
        }, "rangelink.delete-threaded-comment",
        """{"sheetName":"Review","cellAddress":"B2"}""");
    }

    [Fact]
    public async Task QueryTable_AllActionsLifecycleMetadataValidationAndCleanup_ViaMcp()
    {
        const string sourcePath = @"C:\adapter\orders.csv";
        var invalid = await _fixture.CallToolAsync(
            "querytable",
            QueryTableCreateTextArgs(sourcePath, ",,"),
            Failure(
                "querytable.create-text",
                "delimiter must be a single character."),
            "querytable.create-text",
            """{"queryTableName":"CsvImport","sourcePath":"C:\\adapter\\orders.csv","sheetName":"Imports","destinationAddress":"B2","delimiter":",,","textQualifier":"double-quote","encoding":65001,"hasHeaders":true}""");
        AssertRequest(invalid.Request, "querytable.create-text");
        AssertFailure(invalid.JsonResult);

        await AssertSuccessAsync(
            "querytable",
            QueryTableCreateTextArgs(sourcePath, ","),
            "querytable.create-text",
            """{"queryTableName":"CsvImport","sourcePath":"C:\\adapter\\orders.csv","sheetName":"Imports","destinationAddress":"B2","delimiter":",","textQualifier":"double-quote","encoding":65001,"hasHeaders":true}""",
            args =>
            {
                Assert.Equal(sourcePath, args.GetProperty("sourcePath").GetString());
                Assert.Equal(",", args.GetProperty("delimiter").GetString());
                Assert.Equal(
                    "double-quote",
                    args.GetProperty("textQualifier").GetString());
                Assert.Equal(65001, args.GetProperty("encoding").GetInt32());
                Assert.True(args.GetProperty("hasHeaders").GetBoolean());
            });

        var listJson = await CallAsync(
            "querytable",
            new()
            {
                ["action"] = "list",
                ["session_id"] = SessionId
            },
            "querytable.list",
            null,
            """
            {
              "success": true,
              "queryTables": [{
                "name": "CsvImport",
                "sheetName": "Imports",
                "destination": "B2",
                "sourceType": "text"
              }]
            }
            """);
        using (var list = JsonDocument.Parse(listJson))
        {
            var item = Assert.Single(
                list.RootElement.GetProperty("queryTables").EnumerateArray());
            Assert.Equal("CsvImport", item.GetProperty("name").GetString());
            Assert.Equal("Imports", item.GetProperty("sheetName").GetString());
            Assert.Equal("B2", item.GetProperty("destination").GetString());
            Assert.Equal("text", item.GetProperty("sourceType").GetString());
        }

        var viewJson = await CallAsync(
            "querytable",
            QueryTableIdentityArgs("view", "CsvImport"),
            "querytable.view",
            """{"sheetName":"Imports","queryTableName":"CsvImport"}""",
            """
            {
              "success": true,
              "delimiter": ",",
              "encoding": 65001,
              "sourceType": "text",
              "connection": "TEXT;(redacted)"
            }
            """);
        using (var view = JsonDocument.Parse(viewJson))
        {
            var root = view.RootElement;
            Assert.Equal(",", root.GetProperty("delimiter").GetString());
            Assert.Equal(65001, root.GetProperty("encoding").GetInt32());
            Assert.Equal("text", root.GetProperty("sourceType").GetString());
            Assert.Contains(
                "(redacted)",
                root.GetProperty("connection").GetString(),
                StringComparison.Ordinal);
        }

        var invalidProperties = await _fixture.CallToolAsync(
            "querytable",
            new()
            {
                ["action"] = "set-properties",
                ["session_id"] = SessionId,
                ["sheet_name"] = "Imports",
                ["query_table_name"] = "CsvImport",
                ["refresh_period"] = -1
            },
            Failure(
                "querytable.set-properties",
                "refresh_period must be non-negative."),
            "querytable.set-properties",
            """{"sheetName":"Imports","queryTableName":"CsvImport","refreshPeriod":-1}""");
        AssertRequest(invalidProperties.Request, "querytable.set-properties");
        AssertFailure(invalidProperties.JsonResult);

        await AssertSuccessAsync("querytable", new()
        {
            ["action"] = "set-properties",
            ["session_id"] = SessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = "CsvImport",
            ["background_query"] = false,
            ["refresh_on_file_open"] = true,
            ["refresh_period"] = 15,
            ["adjust_column_width"] = false,
            ["preserve_formatting"] = true
        }, "querytable.set-properties",
        """{"sheetName":"Imports","queryTableName":"CsvImport","backgroundQuery":false,"refreshOnFileOpen":true,"refreshPeriod":15,"adjustColumnWidth":false,"preserveFormatting":true}""", args =>
        {
            Assert.False(args.GetProperty("backgroundQuery").GetBoolean());
            Assert.True(args.GetProperty("refreshOnFileOpen").GetBoolean());
            Assert.Equal(15, args.GetProperty("refreshPeriod").GetInt32());
            Assert.False(args.GetProperty("adjustColumnWidth").GetBoolean());
            Assert.True(args.GetProperty("preserveFormatting").GetBoolean());
        });
        await AssertSuccessAsync(
            "querytable",
            QueryTableIdentityArgs("refresh", "CsvImport"),
            "querytable.refresh",
            """{"sheetName":"Imports","queryTableName":"CsvImport"}""");

        var statusJson = await CallAsync(
            "querytable",
            QueryTableIdentityArgs("get-refresh-status", "CsvImport"),
            "querytable.get-refresh-status",
            """{"sheetName":"Imports","queryTableName":"CsvImport"}""",
            """{"success":true,"supportsRefreshStatus":true,"isRefreshing":false}""");
        using (var status = JsonDocument.Parse(statusJson))
        {
            Assert.True(status.RootElement
                .GetProperty("supportsRefreshStatus")
                .GetBoolean());
            Assert.False(status.RootElement
                .GetProperty("isRefreshing")
                .GetBoolean());
        }

        var cancelJson = await CallAsync(
            "querytable",
            QueryTableIdentityArgs("cancel-refresh", "CsvImport"),
            "querytable.cancel-refresh",
            """{"sheetName":"Imports","queryTableName":"CsvImport"}""",
            """{"success":true,"supportsCancellation":true,"wasRefreshing":false,"cancelled":false}""");
        using (var cancel = JsonDocument.Parse(cancelJson))
        {
            Assert.True(cancel.RootElement
                .GetProperty("supportsCancellation")
                .GetBoolean());
            Assert.False(cancel.RootElement.GetProperty("wasRefreshing").GetBoolean());
            Assert.False(cancel.RootElement.GetProperty("cancelled").GetBoolean());
        }

        await AssertSuccessAsync(
            "querytable",
            QueryTableIdentityArgs("delete", "CsvImport"),
            "querytable.delete",
            """{"sheetName":"Imports","queryTableName":"CsvImport"}""");
        await AssertSuccessAsync("querytable", new()
        {
            ["action"] = "create-web",
            ["session_id"] = SessionId,
            ["query_table_name"] = "HtmlImport",
            ["url"] = "file:///C:/adapter/rates.html",
            ["sheet_name"] = "Imports",
            ["destination_address"] = "A1",
            ["selection_type"] = "specified-tables",
            ["web_tables"] = "1",
            ["formatting"] = "none"
        }, "querytable.create-web",
        """{"queryTableName":"HtmlImport","url":"file:///C:/adapter/rates.html","sheetName":"Imports","destinationAddress":"A1","selectionType":"specified-tables","webTables":"1","formatting":"none"}""", args =>
        {
            Assert.Equal(
                "specified-tables",
                args.GetProperty("selectionType").GetString());
            Assert.Equal("1", args.GetProperty("webTables").GetString());
            Assert.Equal("none", args.GetProperty("formatting").GetString());
        });

        var webViewJson = await CallAsync(
            "querytable",
            QueryTableIdentityArgs("view", "HtmlImport"),
            "querytable.view",
            """{"sheetName":"Imports","queryTableName":"HtmlImport"}""",
            """
            {
              "success": true,
              "sourceType": "web",
              "webSelectionType": "specified-tables",
              "webFormatting": "none",
              "webTables": "1"
            }
            """);
        using var webView = JsonDocument.Parse(webViewJson);
        Assert.Equal(
            "web",
            webView.RootElement.GetProperty("sourceType").GetString());
        Assert.Equal(
            "specified-tables",
            webView.RootElement.GetProperty("webSelectionType").GetString());
        Assert.Equal(
            "none",
            webView.RootElement.GetProperty("webFormatting").GetString());
        Assert.Equal(
            "1",
            webView.RootElement.GetProperty("webTables").GetString());
    }

    [Fact]
    public async Task ConnectionRefreshControl_AllActionsMetadataValidationAndCleanup_ViaMcp()
    {
        const string connectionName = "ProductsConnection";
        const string connectionString =
            "OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\\adapter\\Source.xlsx";
        await AssertSuccessAsync("connection", new()
        {
            ["action"] = "create",
            ["session_id"] = SessionId,
            ["connection_name"] = connectionName,
            ["connection_string"] = connectionString,
            ["command_text"] = "SELECT * FROM [Sheet1$]"
        }, "connection.create",
        """{"connectionName":"ProductsConnection","connectionString":"OLEDB;Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\\adapter\\Source.xlsx","commandText":"SELECT * FROM [Sheet1$]"}""", args =>
        {
            Assert.Equal(
                connectionName,
                args.GetProperty("connectionName").GetString());
            Assert.Equal(
                connectionString,
                args.GetProperty("connectionString").GetString());
            Assert.Equal(
                "SELECT * FROM [Sheet1$]",
                args.GetProperty("commandText").GetString());
        });

        var statusJson = await CallAsync(
            "connection",
            ConnectionArgs("get-refresh-status", connectionName),
            "connection.get-refresh-status",
            """{"connectionName":"ProductsConnection"}""",
            """{"success":true,"supportsRefreshStatus":true,"isRefreshing":false}""");
        using (var status = JsonDocument.Parse(statusJson))
        {
            Assert.True(status.RootElement
                .GetProperty("supportsRefreshStatus")
                .GetBoolean());
            Assert.False(status.RootElement
                .GetProperty("isRefreshing")
                .GetBoolean());
        }

        var cancelJson = await CallAsync(
            "connection",
            ConnectionArgs("cancel-refresh", connectionName),
            "connection.cancel-refresh",
            """{"connectionName":"ProductsConnection"}""",
            """{"success":true,"supportsCancellation":true,"wasRefreshing":false,"cancelled":false}""");
        using (var cancel = JsonDocument.Parse(cancelJson))
        {
            Assert.True(cancel.RootElement
                .GetProperty("supportsCancellation")
                .GetBoolean());
            Assert.False(cancel.RootElement.GetProperty("wasRefreshing").GetBoolean());
            Assert.False(cancel.RootElement.GetProperty("cancelled").GetBoolean());
        }

        var missing = await _fixture.CallToolAsync(
            "connection",
            ConnectionArgs("get-refresh-status", "MissingConnection"),
            Failure(
                "connection.get-refresh-status",
                "Connection 'MissingConnection' was not found."),
            "connection.get-refresh-status",
            """{"connectionName":"MissingConnection"}""");
        AssertRequest(missing.Request, "connection.get-refresh-status");
        AssertFailure(missing.JsonResult);

        await AssertSuccessAsync(
            "connection",
            ConnectionArgs("delete", connectionName),
            "connection.delete",
            """{"connectionName":"ProductsConnection"}""");
    }

    private static Dictionary<string, object?> CommentArgs(
        string action,
        string text) => new()
        {
            ["action"] = action,
            ["session_id"] = SessionId,
            ["sheet_name"] = "Review",
            ["cell_address"] = "B2",
            ["text"] = text
        };

    private static Dictionary<string, object?> QueryTableCreateTextArgs(
        string sourcePath,
        string delimiter) => new()
        {
            ["action"] = "create-text",
            ["session_id"] = SessionId,
            ["query_table_name"] = "CsvImport",
            ["source_path"] = sourcePath,
            ["sheet_name"] = "Imports",
            ["destination_address"] = "B2",
            ["delimiter"] = delimiter,
            ["text_qualifier"] = "double-quote",
            ["encoding"] = 65001,
            ["has_headers"] = true
        };

    private static Dictionary<string, object?> QueryTableIdentityArgs(
        string action,
        string name) => new()
        {
            ["action"] = action,
            ["session_id"] = SessionId,
            ["sheet_name"] = "Imports",
            ["query_table_name"] = name
        };

    private static Dictionary<string, object?> ConnectionArgs(
        string action,
        string name) => new()
        {
            ["action"] = action,
            ["session_id"] = SessionId,
            ["connection_name"] = name
        };

    private static ServiceResponse Failure(string command, string message) =>
        new()
        {
            Success = false,
            Command = command,
            SessionId = SessionId,
            ErrorMessage = message,
            ExceptionType = nameof(ArgumentException)
        };

    private static void AssertFailure(string json)
    {
        using var document = JsonDocument.Parse(json);
        Assert.False(document.RootElement.GetProperty("success").GetBoolean());
        Assert.False(string.IsNullOrWhiteSpace(
            document.RootElement.GetProperty("errorMessage").GetString()));
    }

    private static void AssertRequest(ServiceRequest request, string command)
    {
        Assert.Equal(command, request.Command);
        Assert.Equal(SessionId, request.SessionId);
    }

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
        if (assertArgs is not null)
        {
            Assert.NotNull(call.Request.Args);
            using var args = JsonDocument.Parse(call.Request.Args);
            assertArgs(args.RootElement);
        }

        return call.JsonResult;
    }
}
