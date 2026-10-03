using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Layer", "Service")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Slow")]
public class PersistentServiceTableDaxTests(
    PersistentServiceDataModelFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceDataModelFixture>
{
    private readonly ITableCommands _tables = fixture.CreateCommands<ITableCommands>();
    private readonly IDataModelCommands _model = fixture.CreateCommands<IDataModelCommands>();
    private readonly IDataModelRelCommands _relationships = fixture.CreateCommands<IDataModelRelCommands>();
    private readonly Dictionary<string, string> _sheets = new(StringComparer.Ordinal);
    private const string SalesQuery = "EVALUATE 'SalesTable' ORDER BY 'SalesTable'[SalesID]";
    private static readonly string[] SalesColumns =
        ["SalesID", "Date", "CustomerID", "ProductID", "Amount", "Quantity"];
    private static readonly object[][] SalesRows =
    [
        [1, new DateTime(2024, 1, 15).ToOADate(), 101, 1001, 150, 2],
        [2, new DateTime(2024, 1, 20).ToOADate(), 102, 1002, 250, 3],
        [3, new DateTime(2024, 2, 10).ToOADate(), 101, 1003, 175, 1],
        [4, new DateTime(2024, 2, 15).ToOADate(), 103, 1001, 300, 4],
        [5, new DateTime(2024, 3, 5).ToOADate(), 102, 1002, 125, 2],
        [6, new DateTime(2024, 3, 10).ToOADate(), 104, 1003, 450, 5],
        [7, new DateTime(2024, 4, 12).ToOADate(), 101, 1001, 200, 2],
        [8, new DateTime(2024, 4, 18).ToOADate(), 103, 1002, 350, 4],
        [9, new DateTime(2024, 5, 8).ToOADate(), 105, 1003, 275, 3],
        [10, new DateTime(2024, 5, 22).ToOADate(), 102, 1001, 180, 2]
    ];

    [Fact]
    public void CreateFromDax_SimpleEvaluateQuery_CreatesTable()
    {
        var name = Create(SalesQuery);
        AssertTable(name, SalesQuery, SalesColumns, SalesRows);
    }

    [Fact]
    public void CreateFromDax_SummarizeQuery_CreatesAggregatedTable()
    {
        const string query = """
            EVALUATE SUMMARIZE('SalesTable', 'SalesTable'[CustomerID],
                "TotalAmount", SUM('SalesTable'[Amount])) ORDER BY 'SalesTable'[CustomerID]
            """;
        var name = Create(query);
        AssertTable(name, query, ["CustomerID", "TotalAmount"],
            [[101, 525], [102, 555], [103, 650], [104, 450], [105, 275]]);
    }

    [Fact]
    public void CreateFromDax_FilterQuery_CreatesFilteredTable()
    {
        const string query =
            "EVALUATE FILTER('SalesTable', 'SalesTable'[Amount] > 250) ORDER BY 'SalesTable'[SalesID]";
        var name = Create(query);
        AssertTable(name, query, SalesColumns, [SalesRows[3], SalesRows[5], SalesRows[7], SalesRows[8]]);
    }

    [Fact]
    public void CreateFromDax_CustomTargetCell_PlacesTableCorrectly()
    {
        const string query = "EVALUATE 'CustomersTable' ORDER BY 'CustomersTable'[CustomerID]";
        var name = Create(query, "C5");
        AssertTable(name, query, ["CustomerID", "Name", "Region", "Country"],
            [[101, "Acme Corp", "North", "USA"], [102, "Beta Inc", "South", "USA"],
             [103, "Gamma LLC", "East", "Canada"], [104, "Delta Co", "West", "Canada"],
             [105, "Epsilon Ltd", "North", "UK"]], row: 5, column: 3);
        Assert.All(RequireSuccess(_commands.GetValues(_fixture.BatchToken, _sheets[name], "A1:B10")).Values,
            cells => Assert.All(cells, Assert.Null));
    }

    [Fact]
    public void CreateFromDax_WithLocalizedTableName_CreatesReadableTable()
    {
        var name = Create(SalesQuery, name: $"表{Guid.NewGuid():N}");
        AssertTable(name, SalesQuery, SalesColumns, SalesRows);
    }

    [Fact]
    public void UpdateDax_ExistingDaxTable_UpdatesQuery()
    {
        var name = Create(SalesQuery);
        AssertTable(name, SalesQuery, SalesColumns, SalesRows);
        var guard = CaptureModelState();
        const string updated =
            "EVALUATE FILTER('SalesTable', 'SalesTable'[CustomerID] = 101) ORDER BY 'SalesTable'[SalesID]";
        RequireSuccess(_tables.UpdateDax(_fixture.BatchToken, name, updated));
        AssertTable(name, updated, SalesColumns, [SalesRows[0], SalesRows[2], SalesRows[6]]);
        Assert.Equal(guard, CaptureModelState());
    }

    [Fact]
    public void UpdateDax_NonDaxTable_ThrowsError()
    {
        var before = CaptureModelState();
        var regular = JsonSerializer.Serialize(RequireSuccess(_tables.Read(_fixture.BatchToken, "SalesTable")).Table);
        var error = Assert.Throws<InvalidOperationException>(() =>
            _tables.UpdateDax(_fixture.BatchToken, "SalesTable", "EVALUATE 'ProductsTable'"));
        Assert.Contains("Only DAX-backed tables", error.Message, StringComparison.Ordinal);
        Assert.Equal(regular, JsonSerializer.Serialize(
            RequireSuccess(_tables.Read(_fixture.BatchToken, "SalesTable")).Table));
        Assert.Equal(before, CaptureModelState());
    }

    [Fact]
    public void UpdateDax_InvalidDax_ThrowsError()
    {
        var name = Create(SalesQuery);
        AssertTable(name, SalesQuery, SalesColumns, SalesRows);
        var before = CaptureModelState();
        const string invalid = "EVALUATE INVALID_SYNTAX()";
        var error = Assert.Throws<InvalidOperationException>(() =>
            _tables.UpdateDax(_fixture.BatchToken, name, invalid));
        Assert.Contains("table.update-dax failed", error.Message, StringComparison.Ordinal);
        Assert.Contains("ComInterop/", error.Message, StringComparison.Ordinal);
        // The command is stored before execution; failure does not promise command rollback.
        AssertTable(name, invalid, SalesColumns, SalesRows);
        Assert.Equal(before, CaptureModelState());
        RequireSuccess(_tables.UpdateDax(_fixture.BatchToken, name, SalesQuery));
        AssertTable(name, SalesQuery, SalesColumns, SalesRows);
        Assert.Equal(before, CaptureModelState());
    }

    [Fact]
    public void GetDax_DaxBackedTable_ReturnsQueryInfo()
    {
        const string query =
            "EVALUATE TOPN(3, 'SalesTable', 'SalesTable'[Amount], DESC) ORDER BY 'SalesTable'[Amount] DESC";
        var name = Create(query);
        AssertTable(name, query, SalesColumns, [SalesRows[5], SalesRows[7], SalesRows[3]]);
    }

    [Fact]
    public void GetDax_RegularTable_ReturnsNoDaxConnection()
    {
        var before = CaptureModelState();
        var result = RequireSuccess(_tables.GetDax(_fixture.BatchToken, "SalesTable"));
        Assert.Equal("SalesTable", result.TableName);
        Assert.False(result.HasDaxConnection);
        Assert.Null(result.DaxQuery);
        Assert.Null(result.ModelConnectionName);
        Assert.Equal(before, CaptureModelState());
    }

    [Fact]
    public void GetDax_NonExistentTable_ThrowsError()
    {
        var guard = Create(SalesQuery);
        AssertTable(guard, SalesQuery, SalesColumns, SalesRows);
        var before = CaptureModelState();
        var error = Assert.Throws<InvalidOperationException>(() =>
            _tables.GetDax(_fixture.BatchToken, "NonExistentTable_12345"));
        Assert.Contains("not found", error.Message, StringComparison.OrdinalIgnoreCase);
        AssertTable(guard, SalesQuery, SalesColumns, SalesRows);
        Assert.Equal(before, CaptureModelState());
    }

    [Fact]
    public void CreateFromDax_NullSheetName_ThrowsArgumentException() =>
        AssertRejectedCreate(null, "TestTable", SalesQuery, "sheetName");

    [Fact]
    public void CreateFromDax_NullTableName_ThrowsArgumentException() =>
        AssertRejectedCreate("Sheet1", null, SalesQuery, "tableName");

    [Fact]
    public void CreateFromDax_NullDaxQuery_ThrowsArgumentException() =>
        AssertRejectedCreate("Sheet1", "TestTable", null, "daxQuery");

    [Fact]
    public void UpdateDax_NullDaxQuery_ThrowsArgumentException()
    {
        var name = Create(SalesQuery);
        AssertTable(name, SalesQuery, SalesColumns, SalesRows);
        var before = CaptureModelState();
        var error = Assert.Throws<ArgumentException>(() =>
            _tables.UpdateDax(_fixture.BatchToken, name, null!));
        Assert.Contains("daxQuery", error.Message, StringComparison.Ordinal);
        AssertTable(name, SalesQuery, SalesColumns, SalesRows);
        Assert.Equal(before, CaptureModelState());
    }

    private void AssertRejectedCreate(string? sheet, string? name, string? query, string parameter)
    {
        var guard = Create(SalesQuery);
        AssertTable(guard, SalesQuery, SalesColumns, SalesRows);
        var before = CaptureModelState();
        var tables = JsonSerializer.Serialize(RequireSuccess(_tables.List(_fixture.BatchToken)).Tables);
        var connections = CaptureConnectionNames();
        var error = Assert.Throws<ArgumentException>(() =>
            _tables.CreateFromDax(_fixture.BatchToken, sheet!, name!, query!));
        Assert.Contains(parameter, error.Message, StringComparison.Ordinal);
        Assert.Equal(tables, JsonSerializer.Serialize(RequireSuccess(_tables.List(_fixture.BatchToken)).Tables));
        Assert.Equal(connections, CaptureConnectionNames());
        AssertTable(guard, SalesQuery, SalesColumns, SalesRows);
        Assert.Equal(before, CaptureModelState());
    }

    private string Create(string query, string target = "A1", string? name = null)
    {
        name ??= $"DaxTable_{Guid.NewGuid():N}";
        var sheet = _fixture.CreateTestSheet(_fixture.BatchToken);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheet, "J1:K2",
            [["DAX neighbor", "Retained"], [19, 83]]));
        var result = RequireSuccess(_tables.CreateFromDax(_fixture.BatchToken, sheet, name, query, target));
        Assert.Equal(_fixture.WorkbookPath, result.FilePath);
        _fixture.RegisterTableForCleanup(name);
        _sheets.Add(name, sheet);
        return name;
    }

    private void AssertTable(
        string name, string query, string[] columns, object[][] rows, int row = 1, int column = 1)
    {
        var batch = _fixture.BatchToken;
        var data = RequireSuccess(_tables.GetData(batch, name, visibleOnly: false));
        Assert.Equal(name, data.TableName);
        Assert.Equal(columns, data.Headers);
        Assert.Equal(columns.Length, data.ColumnCount);
        Assert.Equal(rows.Length, data.RowCount);
        PowerQueryStateAssertions.AssertRows(rows, data.Data);
        var info = RequireSuccess(_tables.Read(batch, name)).Table;
        Assert.NotNull(info);
        Assert.Equal(name, info.Name);
        Assert.Equal(_sheets[name], info.SheetName);
        Assert.Equal(rows.Length, info.RowCount);
        Assert.Equal(columns.Length, info.ColumnCount);
        Assert.Equal(columns, info.Columns);
        Assert.True(info.HasHeaders);
        Assert.False(info.ShowTotals);
        var listed = Assert.Single(RequireSuccess(_tables.List(batch)).Tables, table => table.Name == name);
        Assert.Equal(JsonSerializer.Serialize(info), JsonSerializer.Serialize(listed));
        var dax = RequireSuccess(_tables.GetDax(batch, name));
        Assert.True(dax.HasDaxConnection);
        Assert.Equal(name, dax.TableName);
        Assert.Equal(query, dax.DaxQuery);
        Assert.False(string.IsNullOrWhiteSpace(dax.ModelConnectionName));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            Excel.Range? range = null;
            Excel.Range? rangeRows = null;
            Excel.Range? rangeColumns = null;
            Excel.TableObject? tableObject = null;
            Excel.WorkbookConnection? connection = null;
            Excel.ModelConnection? modelConnection = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item[_sheets[name]];
                tables = sheet.ListObjects;
                table = tables.Item[name];
                Assert.Equal(Excel.XlListObjectSourceType.xlSrcModel, table.SourceType);
                range = table.Range;
                rangeRows = range.Rows;
                rangeColumns = range.Columns;
                Assert.Equal(row, range.Row);
                Assert.Equal(column, range.Column);
                Assert.Equal(rows.Length + 1, rangeRows.Count);
                Assert.Equal(columns.Length, rangeColumns.Count);
                Assert.Equal(range.Address, info.Range);
                tableObject = table.TableObject;
                connection = tableObject.WorkbookConnection;
                modelConnection = connection.ModelConnection;
                Assert.Equal(dax.ModelConnectionName, connection.Name);
                Assert.Equal(Excel.XlCmdType.xlCmdDAX, modelConnection.CommandType);
                Assert.Equal(query, Convert.ToString(modelConnection.CommandText, CultureInfo.InvariantCulture));
            }
            finally
            {
                ComUtilities.Release(ref modelConnection);
                ComUtilities.Release(ref connection);
                ComUtilities.Release(ref tableObject);
                ComUtilities.Release(ref rangeColumns);
                ComUtilities.Release(ref rangeRows);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        PowerQueryStateAssertions.AssertRows([["DAX neighbor", "Retained"], [19, 83]],
            RequireSuccess(_commands.GetValues(batch, _sheets[name], "J1:K2")).Values);
    }

    private string CaptureModelState()
    {
        var batch = _fixture.BatchToken;
        var tables = RequireSuccess(_model.ListTables(batch)).Tables;
        var measures = RequireSuccess(_model.ListMeasures(batch)).Measures;
        Assert.Equal(5, tables.Count);
        Assert.Equal(6, measures.Count);
        PowerQueryStateAssertions.AssertRows(SalesRows,
            RequireSuccess(_tables.GetData(batch, "SalesTable", visibleOnly: false)).Data);
        var total = RequireSuccess(_model.Evaluate(batch, "EVALUATE ROW(\"Total\", [Total Sales])"));
        PowerQueryStateAssertions.AssertRows([[2455]], total.Rows);
        return JsonSerializer.Serialize(new
        {
            Tables = tables,
            Columns = tables.Select(table => RequireSuccess(_model.ListColumns(batch, table.Name))).ToList(),
            Measures = measures.Select(measure => RequireSuccess(_model.Read(batch, measure.Name))).ToList(),
            Relationships = RequireSuccess(_relationships.ListRelationships(batch)).Relationships,
            Rows = tables.Select(table =>
            {
                var result = RequireSuccess(_model.Evaluate(batch, $"EVALUATE '{table.Name}'"));
                Assert.NotEmpty(result.Rows);
                Assert.Equal(result.RowCount, result.Rows.Count);
                Assert.All(result.Rows, values => Assert.Equal(result.ColumnCount, values.Count));
                return new { table.Name, result.Columns, result.Rows };
            }).ToList(),
            Cells = RequireSuccess(_commands.GetValues(batch, "SalesData", "A1:F11")).Values
        });
    }

    private string CaptureConnectionNames() =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Connections? connections = null;
            try
            {
                connections = context.Book.Connections;
                var names = new List<string>();
                for (var index = 1; index <= connections.Count; index++)
                {
                    Excel.WorkbookConnection? connection = null;
                    try
                    {
                        connection = connections.Item(index);
                        names.Add(connection.Name);
                    }
                    finally { ComUtilities.Release(ref connection); }
                }
                return JsonSerializer.Serialize(names);
            }
            finally { ComUtilities.Release(ref connections); }
        });
}
