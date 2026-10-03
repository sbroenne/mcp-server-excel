using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop;
using Excel = Microsoft.Office.Interop.Excel;
using Xunit;

namespace Sbroenne.ExcelMcp.Service.Tests;

[Collection("ServiceWorkflow")]
[Trait("Category", "Integration")]
[Trait("Layer", "Service")]
[Trait("Feature", "PowerQuery")]
[Trait("RequiresExcel", "true")]
public sealed class PersistentServicePowerQueryWorksheetWorkflowTests(
    PersistentServiceWorkbookFixture fixture) :
    PersistentServiceWorkbookTestBase(fixture),
    IClassFixture<PersistentServiceWorkbookFixture>
{
    private readonly IPowerQueryCommands _queries =
        ServiceCommandProxy.Create<IPowerQueryCommands>(fixture);
    private readonly ISheetCommands _sheets =
        ServiceCommandProxy.Create<ISheetCommands>(fixture);

    [Fact]
    public void Import_QueryReferencingAnotherQuery_LoadsDataSuccessfully()
    {
        var sourceQuery = UniqueName("SourceQuery");
        var derivedQuery = UniqueName("DerivedQuery");
        const string sourceMCode = """
            let
                Source = #table(
                    {"ProductID", "ProductName", "Price"},
                    {
                        {1, "Widget", 10.99},
                        {2, "Gadget", 25.50},
                        {3, "Doohickey", 15.75}
                    }
                )
            in
                Source
            """;
        var derivedMCode = $$"""
            let
                Source = {{sourceQuery}},
                FilteredRows = Table.SelectRows(Source, each [Price] > 15)
            in
                FilteredRows
            """;
        var batch = _fixture.BatchToken;

        CreateTableQuery(sourceQuery, sourceMCode, sourceQuery);
        CreateTableQuery(derivedQuery, derivedMCode, derivedQuery);

        var list = RequireSuccess(_queries.List(batch));
        Assert.True(list.Success);
        Assert.Equal(2, list.Queries.Count);
        Assert.Contains(list.Queries, query => query.Name == sourceQuery);
        Assert.Contains(list.Queries, query => query.Name == derivedQuery);
        var view = RequireSuccess(_queries.View(batch, derivedQuery));
        Assert.True(view.Success);
        Assert.Contains(sourceQuery, view.MCode);
        Assert.Contains("Table.SelectRows", view.MCode);
        PowerQueryStateAssertions.AssertStored(_fixture, sourceQuery, sourceMCode,
            PowerQueryLoadMode.LoadToTable, sourceQuery, ["ProductID", "ProductName", "Price"],
            [[1, "Widget", 10.99m], [2, "Gadget", 25.50m], [3, "Doohickey", 15.75m]]);
        PowerQueryStateAssertions.AssertStored(_fixture, derivedQuery, derivedMCode,
            PowerQueryLoadMode.LoadToTable, derivedQuery, ["ProductID", "ProductName", "Price"],
            [[2, "Gadget", 25.50m], [3, "Doohickey", 15.75m]]);
        const string updatedSource = """
            let Source = #table({"ProductID", "ProductName", "Price"},
                {{1, "Widget", 35}, {2, "Gadget", 10}, {3, "Doohickey", 7.25}, {4, "NewItem", 52.5}})
            in Source
            """;
        RequireSuccess(_queries.Update(batch, sourceQuery, updatedSource, refresh: false));
        PowerQueryStateAssertions.AssertStored(_fixture, sourceQuery, updatedSource,
            PowerQueryLoadMode.LoadToTable, sourceQuery, ["ProductID", "ProductName", "Price"],
            [[1, "Widget", 10.99m], [2, "Gadget", 25.50m], [3, "Doohickey", 15.75m]]);
        var sourceRefresh = RequireSuccess(_queries.Refresh(
            batch,
            sourceQuery,
            TimeSpan.FromMinutes(5)));
        Assert.True(
            sourceRefresh.Success,
            $"Source query refresh failed: {sourceRefresh.ErrorMessage}");
        PowerQueryStateAssertions.AssertStored(_fixture, sourceQuery, updatedSource,
            PowerQueryLoadMode.LoadToTable, sourceQuery, ["ProductID", "ProductName", "Price"],
            [[1, "Widget", 35], [2, "Gadget", 10], [3, "Doohickey", 7.25m], [4, "NewItem", 52.5m]]);
        PowerQueryStateAssertions.AssertStored(_fixture, derivedQuery, derivedMCode,
            PowerQueryLoadMode.LoadToTable, derivedQuery, ["ProductID", "ProductName", "Price"],
            [[2, "Gadget", 25.50m], [3, "Doohickey", 15.75m]]);
        var derivedRefresh = RequireSuccess(_queries.Refresh(
            batch,
            derivedQuery,
            TimeSpan.FromMinutes(5)));
        Assert.True(
            derivedRefresh.Success,
            $"Derived query refresh failed: {derivedRefresh.ErrorMessage}");
        PowerQueryStateAssertions.AssertStored(_fixture, derivedQuery, derivedMCode,
            PowerQueryLoadMode.LoadToTable, derivedQuery, ["ProductID", "ProductName", "Price"],
            [[1, "Widget", 35], [4, "NewItem", 52.5m]]);
        PowerQueryStateAssertions.AssertStored(_fixture, sourceQuery, updatedSource,
            PowerQueryLoadMode.LoadToTable, sourceQuery, ["ProductID", "ProductName", "Price"],
            [[1, "Widget", 35], [2, "Gadget", 10], [3, "Doohickey", 7.25m], [4, "NewItem", 52.5m]]);
    }

    [Fact]
    public void Update_QueryLoadedToSheet_PreservesLoadConfiguration()
    {
        var queryName = UniqueName("LoadedQuery");
        var sheetName = UniqueName("DataSheet");
        var batch = _fixture.BatchToken;
        CreateTableQuery(queryName, InitialTableMCode, sheetName);
        AssertInitial(queryName, sheetName);

        var before = RequireSuccess(_queries.GetLoadConfig(batch, queryName));
        Assert.True(before.Success, "GetLoadConfig before update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, before.LoadMode);
        Assert.Equal(sheetName, before.TargetSheet);

        RequireSuccess(_queries.Update(batch, queryName, UpdatedTableMCode));
        AssertUpdated(queryName, sheetName);

        var after = RequireSuccess(_queries.GetLoadConfig(batch, queryName));
        Assert.True(after.Success, "GetLoadConfig after update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, after.LoadMode);
        Assert.Equal(sheetName, after.TargetSheet);
        var sheets = RequireSuccess(_sheets.List(batch));
        Assert.Contains(sheets.Worksheets, sheet => sheet.Name == sheetName);
    }

    [Fact]
    public void UpdateMCodeThenRefresh_QueryLoadedToSheet_PreservesLoadConfiguration()
    {
        var queryName = UniqueName("LoadedQuery");
        var sheetName = UniqueName("DataSheet");
        var batch = _fixture.BatchToken;
        CreateTableQuery(queryName, InitialTableMCode, sheetName);
        AssertInitial(queryName, sheetName);

        var before = RequireSuccess(_queries.GetLoadConfig(batch, queryName));
        Assert.True(before.Success, "GetLoadConfig before update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, before.LoadMode);
        Assert.Equal(sheetName, before.TargetSheet);

        RequireSuccess(_queries.Update(batch, queryName, UpdatedTableMCode, refresh: false));
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, UpdatedTableMCode,
            PowerQueryLoadMode.LoadToTable, sheetName, ["Column1", "Column2", "Column3"],
            [["Value1", "Value2", "Value3"], ["A", "B", "C"], ["X", "Y", "Z"]]);
        RequireSuccess(_queries.Refresh(batch, queryName, TimeSpan.FromMinutes(5)));
        AssertUpdated(queryName, sheetName);

        var after = RequireSuccess(_queries.GetLoadConfig(batch, queryName));
        Assert.True(after.Success, "GetLoadConfig after update failed");
        Assert.Equal(PowerQueryLoadMode.LoadToTable, after.LoadMode);
        Assert.Equal(sheetName, after.TargetSheet);
        Assert.False(
            string.IsNullOrEmpty(after.TargetSheet),
            "Query should have a target sheet (not be connection-only)");
    }

    [Fact]
    public void Update_QueryColumnStructure_UpdatesWorksheetColumns()
    {
        var queryName = UniqueName("ColumnStructure");
        var sheetName = UniqueName("DataSheet");
        var batch = _fixture.BatchToken;
        const string oneColumnMCode = """
            let
                Source = #table(
                    {"Column1"},
                    {{"Value1"}, {"Value2"}}
                )
            in
                Source
            """;
        const string oneColumnUpdatedMCode = """
            let
                Source = #table(
                    {"Column1"},
                    {{"UpdatedValue1"}, {"UpdatedValue2"}, {"UpdatedValue3"}}
                )
            in
                Source
            """;
        const string twoColumnMCode = """
            let
                Source = #table(
                    {"Column1", "Column2"},
                    {{"A", "B"}, {"C", "D"}, {"E", "F"}}
                )
            in
                Source
            """;
        CreateTableQuery(queryName, oneColumnMCode, sheetName);
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, oneColumnMCode,
            PowerQueryLoadMode.LoadToTable, sheetName, ["Column1"], [["Value1"], ["Value2"]]);

        var initial = RequireSuccess(_commands.GetUsedRange(batch, sheetName));
        Assert.True(initial.Success, initial.ErrorMessage);
        Assert.Equal(1, initial.ColumnCount);
        RequireSuccess(_queries.Update(batch, queryName, oneColumnUpdatedMCode));
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, oneColumnUpdatedMCode,
            PowerQueryLoadMode.LoadToTable, sheetName, ["Column1"],
            [["UpdatedValue1"], ["UpdatedValue2"], ["UpdatedValue3"]]);
        var afterFirstUpdate = RequireSuccess(_commands.GetUsedRange(batch, sheetName));
        Assert.True(afterFirstUpdate.Success, afterFirstUpdate.ErrorMessage);
        Assert.Equal(1, afterFirstUpdate.ColumnCount);

        RequireSuccess(_queries.Update(batch, queryName, twoColumnMCode));
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, twoColumnMCode,
            PowerQueryLoadMode.LoadToTable, sheetName, ["Column1", "Column2"],
            [["A", "B"], ["C", "D"], ["E", "F"]]);

        var afterSecondUpdate = RequireSuccess(_commands.GetUsedRange(batch, sheetName));
        Assert.True(afterSecondUpdate.Success, afterSecondUpdate.ErrorMessage);
        var values = RequireSuccess(_commands.GetValues(
            batch,
            sheetName,
            afterSecondUpdate.RangeAddress));
        Assert.True(values.Success, values.ErrorMessage);
        var headers = values.Values.FirstOrDefault();
        var columnNames = headers is null
            ? "No headers found"
            : string.Join(", ", headers.Select(value => value?.ToString() ?? "null"));
        Assert.True(
            afterSecondUpdate.ColumnCount == 2,
            $"Expected 2 columns but got {afterSecondUpdate.ColumnCount}. " +
            $"Actual columns: [{columnNames}]");
        Assert.True(
            values.ColumnCount == 2,
            $"Expected 2 columns in values but got {values.ColumnCount}. " +
            $"Columns: [{columnNames}]");
    }

    [Fact]
    public void Update_QueryColumnStructureWithDeleteRecreate_NoAccumulation()
    {
        var queryName = UniqueName("AccumulationBug");
        var sheetName = UniqueName("TestSheet");
        var batch = _fixture.BatchToken;
        const string oneColumnMCode = """
            let
                Source = #table({"Column1"}, {{"A"}, {"B"}})
            in
                Source
            """;
        const string twoColumnMCode = """
            let
                Source = #table(
                    {"Column1", "Column2"},
                    {{"X", "Y"}, {"Z", "W"}}
                )
            in
                Source
            """;
        CreateTableQuery(queryName, oneColumnMCode, sheetName);
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, oneColumnMCode,
            PowerQueryLoadMode.LoadToTable, sheetName, ["Column1"], [["A"], ["B"]]);
        var initial = RequireSuccess(_commands.GetUsedRange(batch, sheetName));
        Assert.True(initial.Success);
        Assert.Equal(1, initial.ColumnCount);

        RequireSuccess(_queries.Update(batch, queryName, twoColumnMCode));
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, twoColumnMCode,
            PowerQueryLoadMode.LoadToTable, sheetName, ["Column1", "Column2"], [["X", "Y"], ["Z", "W"]]);
        RequireSuccess(_queries.Delete(batch, queryName));
        _fixture.ForgetPowerQuery(queryName);
        PowerQueryStateAssertions.AssertRemoved(_fixture, queryName);
        Assert.All(RequireSuccess(_commands.GetValues(batch, sheetName, "A1:B3")).Values,
            row => Assert.All(row, Assert.Null));
        RequireSuccess(_queries.Create(
            batch,
            queryName,
            twoColumnMCode,
            PowerQueryLoadMode.ConnectionOnly));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, twoColumnMCode,
            PowerQueryLoadMode.ConnectionOnly, null, ["Column1", "Column2"], [["X", "Y"], ["Z", "W"]]);
        RequireSuccess(_queries.LoadTo(
            batch,
            queryName,
            PowerQueryLoadMode.LoadToTable,
            sheetName,
            "A1"));
        PowerQueryStateAssertions.AssertStored(_fixture, queryName, twoColumnMCode,
            PowerQueryLoadMode.LoadToTable, sheetName, ["Column1", "Column2"], [["X", "Y"], ["Z", "W"]]);
        var refresh = RequireSuccess(_queries.Refresh(
            batch,
            queryName,
            TimeSpan.FromMinutes(5)));
        Assert.True(refresh.Success, $"Refresh failed: {refresh.ErrorMessage}");

        PowerQueryStateAssertions.AssertStored(_fixture, queryName, twoColumnMCode,
            PowerQueryLoadMode.LoadToTable, sheetName, ["Column1", "Column2"], [["X", "Y"], ["Z", "W"]]);
        var usedRange = RequireSuccess(_commands.GetUsedRange(batch, sheetName));
        Assert.True(usedRange.Success);
        var values = RequireSuccess(_commands.GetValues(
            batch,
            sheetName,
            usedRange.RangeAddress));
        Assert.True(values.Success);
        var headers = values.Values.FirstOrDefault();
        var columnNames = headers is null
            ? "No headers found"
            : string.Join(", ", headers.Select(value => value?.ToString() ?? "null"));
        Assert.True(
            usedRange.ColumnCount == 2,
            $"COLUMN ACCUMULATION DETECTED! Expected 2 columns but got " +
            $"{usedRange.ColumnCount}. Actual columns: [{columnNames}].");
    }

    [Theory]
    [InlineData(true, Excel.XlCellInsertionMode.xlInsertDeleteCells)]
    [InlineData(false, Excel.XlCellInsertionMode.xlInsertDeleteCells)]
    [InlineData(true, Excel.XlCellInsertionMode.xlOverwriteCells)]
    [InlineData(false, Excel.XlCellInsertionMode.xlOverwriteCells)]
    public void NativeRefresh_PreserveColumnInfo_ControlsSourceColumnRemoval(
        bool preserveColumnInfo,
        Excel.XlCellInsertionMode refreshStyle)
    {
        var name = UniqueName("NativeSchema");
        var sheetName = UniqueName("NativeSheet");
        CreateTableQuery(name, InitialTableMCode, sheetName);
        AssertInitial(name, sheetName);
        RequireSuccess(_queries.Update(_fixture.BatchToken, name, UpdatedTableMCode, refresh: false));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            Excel.QueryTable? query = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item[sheetName];
                tables = sheet.ListObjects;
                table = tables.Item[1];
                query = table.QueryTable;
                query.PreserveColumnInfo = preserveColumnInfo;
                query.RefreshStyle = refreshStyle;
                Assert.True(query.Refresh(false));
            }
            finally
            {
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        if (preserveColumnInfo)
        {
            AssertUpdated(name, sheetName);
        }
        else
        {
            Assert.Equal("Table=$A$1:$C$3; result=$A$1:$C$3; fields=NewCol1, NewCol2, NewCol3",
                DescribeNativeShape(sheetName));
            var cells = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:C3")).Values;
            Assert.Equal(3, cells.Count);
            Assert.All(cells, row => Assert.Equal(3, row.Count));
            PowerQueryStateAssertions.AssertRows(
                [["Updated1", "Updated2"], ["Data1", "Data2"]],
                cells.Skip(1).Select(row => row.Take(2).ToList()).ToList());
            Assert.All(cells.Skip(1), row => Assert.Null(row[2]));
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Refresh_RemovingSourceColumn_PreservesCalculatedColumnAndNeighbor(bool deferRefresh)
    {
        var name = UniqueName("KeepCalculated");
        var sheetName = UniqueName("CalculatedSheet");
        CreateTableQuery(name, InitialTableMCode, sheetName);
        AssertInitial(name, sheetName);
        RequireSuccess(_commands.SetValues(_fixture.BatchToken, sheetName, "F1:G2",
            [["Neighbor", "Retained"], [91, 47]]));
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            Excel.QueryTable? query = null;
            Excel.ListColumns? columns = null;
            Excel.ListColumn? column = null;
            Excel.Range? data = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item[sheetName];
                tables = sheet.ListObjects;
                table = tables.Item[1];
                query = table.QueryTable;
                query.PreserveColumnInfo = false;
                columns = table.ListColumns;
                column = columns.Add();
                column.Name = "Audit";
                data = column.DataBodyRange;
                data.Formula = "=73";
            }
            finally
            {
                ComUtilities.Release(ref data);
                ComUtilities.Release(ref column);
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        PowerQueryStateAssertions.AssertRows(
            [["Column1", "Column2", "Column3", "Audit"],
             ["Value1", "Value2", "Value3", 73], ["A", "B", "C", 73], ["X", "Y", "Z", 73]],
            RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:D4")).Values);
        RequireSuccess(_queries.Update(_fixture.BatchToken, name, UpdatedTableMCode, refresh: !deferRefresh));
        if (deferRefresh)
        {
            PowerQueryStateAssertions.AssertRows(
                [["Column1", "Column2", "Column3", "Audit"],
                 ["Value1", "Value2", "Value3", 73], ["A", "B", "C", 73], ["X", "Y", "Z", 73]],
                RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:D4")).Values);
            RequireSuccess(_queries.Refresh(_fixture.BatchToken, name, TimeSpan.FromMinutes(5)));
        }
        PowerQueryStateAssertions.AssertStored(_fixture, name, UpdatedTableMCode,
            PowerQueryLoadMode.LoadToTable, sheetName, ["NewCol1", "NewCol2"],
            [["Updated1", "Updated2"], ["Data1", "Data2"]]);
        Assert.Equal("Table=$A$1:$C$3; result=$A$1:$C$3; fields=NewCol1, NewCol2, Audit",
            DescribeNativeShape(sheetName));
        PowerQueryStateAssertions.AssertRows(
            [["NewCol1", "NewCol2", "Audit"], ["Updated1", "Updated2", 73], ["Data1", "Data2", 73]],
            RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A1:C3")).Values);
        Assert.All(RequireSuccess(_commands.GetFormulas(_fixture.BatchToken, sheetName, "C2:C3")).Formulas,
            row => Assert.Equal("=73", Assert.Single(row)));
        Assert.All(RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "D1:D4")).Values,
            row => Assert.All(row, Assert.Null));
        Assert.All(RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "A4:C4")).Values,
            row => Assert.All(row, Assert.Null));
        PowerQueryStateAssertions.AssertRows([["Neighbor", "Retained"], [91, 47]],
            RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheetName, "F1:G2")).Values);
    }

    private void CreateTableQuery(
        string queryName,
        string mCode,
        string sheetName)
    {
        RequireSuccess(_queries.Create(
            _fixture.BatchToken,
            queryName,
            mCode,
            PowerQueryLoadMode.LoadToTable,
            sheetName));
        _fixture.RegisterPowerQueryForCleanup(queryName);
        _fixture.RegisterSheetForCleanup(sheetName);
    }

    private static string UniqueName(string prefix) =>
        $"{prefix}_{Guid.NewGuid():N}"[..Math.Min(prefix.Length + 9, 31)];

    private void AssertInitial(string name, string sheet) =>
        PowerQueryStateAssertions.AssertStored(_fixture, name, InitialTableMCode,
            PowerQueryLoadMode.LoadToTable, sheet, ["Column1", "Column2", "Column3"],
            [["Value1", "Value2", "Value3"], ["A", "B", "C"], ["X", "Y", "Z"]]);

    private void AssertUpdated(string name, string sheet)
    {
        PowerQueryStateAssertions.AssertStored(_fixture, name, UpdatedTableMCode,
            PowerQueryLoadMode.LoadToTable, sheet, ["NewCol1", "NewCol2"],
            [["Updated1", "Updated2"], ["Data1", "Data2"]]);
        var removedColumn = RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, "C1:C4"));
        Assert.True(removedColumn.Values.All(row => row.All(value => value is null)),
            $"Removed source column remains populated. {DescribeNativeShape(sheet)}");
        Assert.Equal("Table=$A$1:$B$3; result=$A$1:$B$3; fields=NewCol1, NewCol2",
            DescribeNativeShape(sheet));
        Assert.All(RequireSuccess(_commands.GetValues(_fixture.BatchToken, sheet, "A4:B4")).Values,
            row => Assert.All(row, Assert.Null));
    }

    private string DescribeNativeShape(string sheetName) =>
        _fixture.ExecuteRawVerification((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            Excel.Range? tableRange = null;
            Excel.QueryTable? query = null;
            Excel.Range? resultRange = null;
            Excel.ListColumns? columns = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets.Item[sheetName];
                tables = sheet.ListObjects;
                table = tables.Item[1];
                tableRange = table.Range;
                query = table.QueryTable;
                resultRange = query.ResultRange;
                columns = table.ListColumns;
                var fieldNames = new List<string>();
                for (var index = 1; index <= columns.Count; index++)
                {
                    Excel.ListColumn? column = null;
                    try
                    {
                        column = columns.Item[index];
                        fieldNames.Add(column.Name);
                    }
                    finally
                    {
                        ComUtilities.Release(ref column);
                    }
                }
                return $"Table={tableRange.Address}; result={resultRange.Address}; fields={string.Join(", ", fieldNames)}";
            }
            finally
            {
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref resultRange);
                ComUtilities.Release(ref query);
                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

    private const string InitialTableMCode = """
        let
            Source = #table(
                {"Column1", "Column2", "Column3"},
                {
                    {"Value1", "Value2", "Value3"},
                    {"A", "B", "C"},
                    {"X", "Y", "Z"}
                }
            )
        in
            Source
        """;

    private const string UpdatedTableMCode = """
        let
            UpdatedSource = #table(
                {"NewCol1", "NewCol2"},
                {{"Updated1", "Updated2"}, {"Data1", "Data2"}}
            )
        in
            UpdatedSource
        """;
}
