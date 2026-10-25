using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class ServiceWorkbookLifecycleTests
{
    [Theory]
    [InlineData(false, true)]
    [InlineData(true, true)]
    [InlineData(false, false)]
    [InlineData(true, false)]
    public async Task LegacyImport_LoadTo_PersistsSettingsAndRefreshesAfterReopen(bool web, bool move)
    {
        await RunWithCleanupAsync(async (service, directory, sessions) =>
        {
            string path = Path.Join(directory, "legacy-move.xlsx");
            string source = Path.Join(directory, web ? "source.html" : "source.csv");
            File.WriteAllText(source, web
                ? "<html><body><table><tr><td>Ignored</td></tr></table>" +
                  "<table><tr><th>Name</th><th>Amount</th></tr><tr><td>Selected</td><td>17</td></tr></table></body></html>"
                : "Name;Amount\r\nSelected;17\r\n");
            string session = await CreateSessionAsync(service, path);
            sessions[session] = 0;
            await CreateSheetAsync(service, session, "Target");
            await Call<OperationResult>(web ? "querytable.create-web" : "querytable.create-text", web
                ? new
                {
                    queryTableName = "LegacyImport",
                    url = new Uri(source).AbsoluteUri,
                    sheetName = "Sheet1",
                    destinationAddress = "A1",
                    selectionType = "specified-tables",
                    webTables = "2"
                }
                : (object)new
                {
                    queryTableName = "LegacyImport",
                    sourcePath = source,
                    sheetName = "Sheet1",
                    destinationAddress = "A1",
                    delimiter = ";"
                });
            string connectionName = Assert.Single((await Call<ConnectionListResult>("connection.list")).Connections).Name;
            await Call<OperationResult>("querytable.set-properties", new
            {
                sheetName = "Sheet1",
                queryTableName = "LegacyImport",
                backgroundQuery = false,
                refreshOnFileOpen = false,
                refreshPeriod = 17,
                adjustColumnWidth = false,
                preserveFormatting = false
            });
            await Call<OperationResult>("connection.set-properties",
                new { connectionName, description = "Retained description" });
            await Call<OperationResult>("range.set-values", new
            {
                sheetName = "Target",
                rangeAddress = "G7",
                overwritePolicy = "allow",
                values = (object?[][])[["Unrelated"]]
            });
            var before = await Call<QueryTableViewResult>("querytable.view",
                new { sheetName = "Sheet1", queryTableName = "LegacyImport" });

            if (move)
            {
                var moved = await Call<OperationResult>("connection.load-to", new { connectionName, sheetName = "Target" });
                Assert.Equal(path, moved.FilePath);
                Assert.Equal("Retained description",
                    Assert.Single((await Call<ConnectionListResult>("connection.list")).Connections).Description);
            }
            await CloseSessionAsync(service, session, save: true);
            sessions.TryRemove(session, out _);
            session = await OpenSessionAsync(service, path);
            sessions[session] = 0;

            var connection = Assert.Single((await Call<ConnectionListResult>("connection.list")).Connections);
            Assert.Equal(connectionName, connection.Name);
            string loadedSheet = move ? "Target" : "Sheet1";
            var query = Assert.Single((await Call<QueryTableListResult>("querytable.list")).QueryTables);
            Assert.Equal(loadedSheet, query.SheetName);
            var after = await Call<QueryTableViewResult>("querytable.view",
                new { sheetName = loadedSheet, queryTableName = query.Name });
            Assert.Equal(before.SourceType, after.SourceType);
            Assert.Equal(before.Connection, after.Connection);
            Assert.Equal(before.Destination, after.Destination);
            Assert.Equal(before.Delimiter, after.Delimiter);
            Assert.Equal(before.Encoding, after.Encoding);
            Assert.Equal(before.WebSelectionType, after.WebSelectionType);
            Assert.Equal(before.WebTables, after.WebTables);
            Assert.Equal(before.WebFormatting, after.WebFormatting);
            Assert.False(after.BackgroundQuery);
            Assert.False(after.RefreshOnFileOpen);
            Assert.Equal(17, after.RefreshPeriod);
            Assert.False(after.AdjustColumnWidth);
            Assert.False(after.PreserveFormatting);
            Assert.Equal("Selected", await ReadMarkerAsync(service, session, "Sheet1", "A2"));
            Assert.Equal("Selected", await ReadMarkerAsync(service, session, loadedSheet, "A2"));
            Assert.Equal("Unrelated", await ReadMarkerAsync(service, session, "Target", "G7"));
            Assert.Equal("Retained description", connection.Description);
            await AssertAmountAsync(loadedSheet, 17);

            File.WriteAllText(source, File.ReadAllText(source)
                .Replace("Selected", "Reloaded", StringComparison.Ordinal)
                .Replace("17", "29", StringComparison.Ordinal));
            await Call<OperationResult>("connection.refresh", new { connectionName });
            Assert.Equal("Reloaded", await ReadMarkerAsync(service, session, loadedSheet, "A2"));
            await AssertAmountAsync(loadedSheet, 29);
            Assert.Equal(move ? "Selected" : "Reloaded", await ReadMarkerAsync(service, session, "Sheet1", "A2"));
            await AssertAmountAsync("Sheet1", move ? 17 : 29);
            Assert.Equal("Unrelated", await ReadMarkerAsync(service, session, "Target", "G7"));

            async Task<T> Call<T>(string command, object? args = null) where T : class
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = command,
                    SessionId = session,
                    Args = args is null ? null : JsonSerializer.Serialize(args, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(response);
                return JsonSerializer.Deserialize<T>(response.Result!, ServiceProtocol.JsonOptions)
                    ?? throw new InvalidOperationException($"Missing result from {command}.");
            }

            async Task AssertAmountAsync(string sheetName, double expected)
            {
                var response = await service.ProcessAsync(new ServiceRequest
                {
                    Command = "range.get-values",
                    SessionId = session,
                    Args = JsonSerializer.Serialize(new { sheetName, rangeAddress = "B2" }, ServiceProtocol.JsonOptions)
                });
                RequireSuccess(response);
                using var result = JsonDocument.Parse(response.Result!);
                Assert.Equal(expected, result.RootElement.GetProperty("values")[0][0].GetDouble());
            }
        });
    }

}
