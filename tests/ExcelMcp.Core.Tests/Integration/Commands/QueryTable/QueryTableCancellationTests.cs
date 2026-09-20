using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.QueryTable;

[Trait("Category", "Integration")]
[Trait("Feature", "QueryTable")]
[Trait("Layer", "Core")]
[Trait("RequiresExcel", "true")]
[Trait("Speed", "Medium")]
public sealed class QueryTableCancellationTests : IClassFixture<TempDirectoryFixture>
{
    private readonly TempDirectoryFixture _fixture;

    public QueryTableCancellationTests(TempDirectoryFixture fixture) => _fixture = fixture;

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Refresh_CancelledByBeforeRefresh_DoesNotReportSuccess(bool powerQuery)
    {
        using var batch = ExcelSession.BeginBatch(_fixture.CreateTestFile());
        var commands = new QueryTableCommands();
        var powerQueryCommands = new PowerQueryCommands(new DataModelCommands());
        const string name = "CancelledImport";
        if (powerQuery)
        {
            powerQueryCommands.Create(batch, name, "#table({\"Value\"}, {{1}})", targetSheet: name);
        }
        else
        {
            var sourcePath = CoreTestHelper.CreateUniqueTestFile(
                nameof(QueryTableCancellationTests), nameof(Refresh_CancelledByBeforeRefresh_DoesNotReportSuccess),
                _fixture.TempDir, ".csv", "Value\n1\n");
            batch.Execute((ctx, ct) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                try
                {
                    sheets = ctx.Book.Worksheets;
                    sheet = (Excel.Worksheet)sheets.Add();
                    sheet.Name = name;
                    return 0;
                }
                finally
                {
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            });
            commands.CreateText(batch, name, sourcePath, name, "A1");
        }

        Excel.QueryTable? queryTable = null;
        string queryTableName = name;
        bool eventRaised = false;
        Excel.RefreshEvents_BeforeRefreshEventHandler handler = (ref bool cancel) =>
        {
            eventRaised = true;
            cancel = true;
        };
        batch.Execute((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.QueryTables? tables = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item[name];
                tables = sheet.QueryTables;
                queryTable = tables.Item(1);
                queryTableName = queryTable.Name;
                queryTable.BeforeRefresh += handler;
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        try
        {
            var error = Assert.ThrowsAny<Exception>(() =>
            {
                if (powerQuery)
                    powerQueryCommands.Refresh(batch, name, TimeSpan.FromSeconds(30));
                else
                    commands.Refresh(batch, name, queryTableName);
            });
            Assert.True(eventRaised);
            Assert.Contains("cancelled", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.False(commands.GetRefreshStatus(batch, name, queryTableName).IsRefreshing);
        }
        finally
        {
            batch.Execute((ctx, ct) =>
            {
                try
                {
                    if (queryTable != null)
                        queryTable.BeforeRefresh -= handler;
                }
                finally
                {
                    ComUtilities.Release(ref queryTable);
                }
                return 0;
            });
        }
    }
}
