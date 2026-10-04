using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

public sealed partial class PersistentServiceTablePreflightTests
{
    [Fact]
    public void Append_WithJsonElementValues_DoesNotThrow()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        RequireSuccess(_rangeCommands.SetValues(
            batch,
            "Data",
            "A1:C2",
            [["Label", "IsActive", "Amount"], ["Initial", true, 1.0]]));
        RequireSuccess(_tableCommands.Create(
            batch, "Data", "DataTable", "A1:C2", true, "TableStyleLight1"));

        var rowsJson =
            """[["NewRow", true, 99.5], ["AnotherRow", false, 0.0]]""";
        var deserializedRows =
            JsonSerializer.Deserialize<List<List<object?>>>(rowsJson)!;

        Assert.IsType<JsonElement>(deserializedRows[0][0]);
        Assert.IsType<JsonElement>(deserializedRows[0][1]);
        Assert.IsType<JsonElement>(deserializedRows[0][2]);

        RequireSuccess(_tableCommands.Append(batch, "DataTable", deserializedRows));

        var info = _tableCommands.Read(batch, "DataTable");
        RequireSuccess(info);
        Assert.Equal(3, info.Table!.RowCount);
        var data = RequireSuccess(_tableCommands.GetData(batch, "DataTable"));
        Assert.Equal(["Label", "IsActive", "Amount"], data.Headers);
        Assert.Equal(3, data.Data.Count);
        Assert.Equal(["Initial", "NewRow", "AnotherRow"], data.Data.Select(row => row[0]));
        Assert.Equal([true, true, false], data.Data.Select(row => Assert.IsType<bool>(row[1])));
        Assert.Equal([1d, 99.5d, 0d], data.Data.Select(row => Convert.ToDouble(row[2], CultureInfo.InvariantCulture)));
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
    }

    [Theory]
    [InlineData("")]
    [InlineData("TableStyleLight1")]
    [InlineData("TableStyleMedium2")]
    public void ListAndRead_WithTableStyle_ReturnsExpectedTableStyle(string tableStyle)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        RequireSuccess(_rangeCommands.SetValues(
            batch,
            sheetName,
            "A1:B2",
            [["Name", "Value"], ["Example", 1]]));
        RequireSuccess(_tableCommands.Create(
            batch,
            sheetName,
            "PlainTable",
            "A1:B2"));

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            try
            {
                sheets = ctx.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[sheetName];
                tables = sheet.ListObjects;
                table = tables["PlainTable"];
                table.TableStyle = tableStyle;
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });

        var list = _tableCommands.List(batch);
        var read = _tableCommands.Read(batch, "PlainTable");

        RequireSuccess(list);
        Assert.Equal(
            tableStyle,
            Assert.Single(list.Tables, table => table.Name == "PlainTable").TableStyle);
        RequireSuccess(read);
        Assert.Equal(tableStyle, read.Table!.TableStyle);
        var data = RequireSuccess(_tableCommands.GetData(batch, "PlainTable"));
        Assert.Equal(["Name", "Value"], data.Headers);
        AssertNamedAmountRow(Assert.Single(data.Data), "Example", 1);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
    }

    [Theory]
    [InlineData("表1")]
    [InlineData("テーブル1")]
    [InlineData("표1")]
    public void Create_WithNonAsciiTableName_CreatesReadableTable(string tableName)
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        RequireSuccess(_rangeCommands.SetValues(
            batch, "Data", "A1:B2", [["Name", "Value"], ["North", 100]]));

        var result = _tableCommands.Create(
            batch, "Data", tableName, "A1:B2", true, "TableStyleLight1");
        RequireSuccess(result);

        var info = _tableCommands.Read(batch, tableName);
        RequireSuccess(info);
        Assert.NotNull(info.Table);
        Assert.Equal(tableName, info.Table.Name);
        Assert.Equal("Data", info.Table.SheetName);
        Assert.Equal("$A$1:$B$2", info.Table.Range);
        var data = RequireSuccess(_tableCommands.GetData(batch, tableName));
        Assert.Equal(["Name", "Value"], data.Headers);
        AssertNamedAmountRow(Assert.Single(data.Data), "North", 100);
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
    }

    private void SetUpSourceState(bool hasHeaders)
    {
        _fixture.CreateNamedTestSheet(_fixture.BatchToken, "Data");
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Microsoft.Office.Interop.Excel.Interior? interior = null;
            Microsoft.Office.Interop.Excel.Font? font = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, "Data")
                    ?? throw new InvalidOperationException("Data not found.");
                range = sheet.Range["A1:C4"];
                range.Formula = new object[,]
                {
                    { hasHeaders ? "Amount" : 10, hasHeaders ? "Calculated" : 20, "Beside" },
                    { 30, "=A2*2", "Keep" },
                    { 50, "=A3*2", "Below" },
                    { 70, 80, "Untouched" }
                };
                range.NumberFormat = "0.0000";
                interior = range.Interior;
                interior.Color = 0x00336699;
                font = range.Font;
                font.Bold = false;
            }
            finally
            {
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private (object[,] Formulas, string[] Formats, int WorkbookCount) GetSourceState(
        string sheetName = "Data", string address = "A1:C4")
    {
        return _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Microsoft.Office.Interop.Excel.Workbooks? workbooks = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"{sheetName} not found.");
                range = sheet.Range[address];
                var formulas = (object[,])range.Formula;
                var formats = new List<string>();
                for (int row = 1; row <= formulas.GetLength(0); row++)
                {
                    for (int column = 1; column <= formulas.GetLength(1); column++)
                    {
                        Excel.Range? cell = null;
                        Microsoft.Office.Interop.Excel.Interior? interior = null;
                        Microsoft.Office.Interop.Excel.Font? font = null;
                        try
                        {
                            cell = (Excel.Range)range[row, column];
                            interior = cell.Interior;
                            font = cell.Font;
                            formats.Add(string.Format(CultureInfo.InvariantCulture,
                                "{0}|{1}|{2}", cell.NumberFormat, interior.Color, font.Bold));
                        }
                        finally
                        {
                            ComUtilities.Release(ref font);
                            ComUtilities.Release(ref interior);
                            ComUtilities.Release(ref cell);
                        }
                    }
                }

                workbooks = ctx.App.Workbooks;
                return (formulas, formats.ToArray(), workbooks.Count);
            }
            finally
            {
                ComUtilities.Release(ref workbooks);
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    [Theory]
    [InlineData("Sales", true)]
    [InlineData("Sales", false)]
    [InlineData("sales", true)]
    [InlineData("sales", false)]
    public void Create_WithExistingDefinedName_CreatesTableAndPreservesName(string tableName, bool hasHeaders)
    {
        var batch = _fixture.BatchToken;
        SetUpSourceState(hasHeaders);
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Microsoft.Office.Interop.Excel.Names? names = null;
            Microsoft.Office.Interop.Excel.Name? name = null;
            try
            {
                names = ctx.Book.Names;
                name = names.Add("Sales", "=Data!$A$1");
            }
            finally
            {
                ComUtilities.Release(ref name);
                ComUtilities.Release(ref names);
            }
        });
        _fixture.RegisterNamedRangeForCleanup("Sales");

        var before = GetSourceState();
        var result = _tableCommands.Create(batch, "Data", tableName, "A1:B2", hasHeaders);
        RequireSuccess(result);
        var after = GetSourceState(address: hasHeaders ? "A1:C4" : "A1:C5");
        Assert.Equal(before.WorkbookCount, after.WorkbookCount);
        Assert.Equal(1, GetWorksheetTableCount("Data"));
        var info = _tableCommands.Read(batch, tableName);
        RequireSuccess(info);
        Assert.NotNull(info.Table);
        Assert.Equal(tableName, info.Table.Name);
        Assert.Equal(hasHeaders ? 1 : 2, info.Table.RowCount);
        Assert.Equal(hasHeaders ? "$A$1:$B$2" : "$A$1:$B$3", info.Table.Range);
        object?[,] expected = hasHeaders
            ? before.Formulas
            : new object?[,]
            {
                { "Column1", "Column2", before.Formulas[1, 3] },
                { before.Formulas[1, 1], before.Formulas[1, 2], before.Formulas[2, 3] },
                { before.Formulas[2, 1], "=A3*2", before.Formulas[3, 3] },
                { before.Formulas[3, 1], "=A4*2", before.Formulas[4, 3] },
                { before.Formulas[4, 1], before.Formulas[4, 2], string.Empty }
            };
        Assert.Equal(expected.Cast<object>(), after.Formulas.Cast<object>());
        if (!hasHeaders) { Assert.Null(GetRangeValues("Data", "C4:C5")[2, 1]); }
        for (var row = 0; row < 4; row++)
        {
            Assert.Equal(before.Formats[row * 3 + 2], after.Formats[row * 3 + 2]);
        }
        for (var row = 2; row < 4; row++)
        {
            for (var column = 0; column < 2; column++)
            {
                Assert.Equal(before.Formats[row * 3 + column],
                    after.Formats[(row + (hasHeaders ? 0 : 1)) * 3 + column]);
            }
        }
        var data = RequireSuccess(_tableCommands.GetData(batch, tableName));
        Assert.Equal(hasHeaders ? ["Amount", "Calculated"] : ["Column1", "Column2"], data.Headers);
        Assert.Equal(hasHeaders ? 1 : 2, data.Data.Count);
        AssertNumericPair(data.Data[^1], 30, 60);
        if (!hasHeaders) { AssertNumericPair(data.Data[0], 10, 20); }
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Microsoft.Office.Interop.Excel.Names? names = null;
            Microsoft.Office.Interop.Excel.Name? name = null;
            try
            {
                names = ctx.Book.Names;
                name = names.Item("Sales");
                Assert.Equal("Sales", name.Name);
                Assert.Equal(hasHeaders ? "=Data!$A$1" : "=Data!$A$2", name.RefersTo);
            }
            finally
            {
                ComUtilities.Release(ref name);
                ComUtilities.Release(ref names);
            }
        });
    }

    [Fact]
    public void ReadAndRename_WithExistingNonAsciiTableName_Succeeds()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        RequireSuccess(_rangeCommands.SetValues(
            batch, "Data", "A1:B2", [["Name", "Value"], ["North", 100]]));
        var result = _tableCommands.Create(
            batch, "Data", "PlainTable", "A1:B2", true, "TableStyleLight1");
        RequireSuccess(result);

        _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? listObjects = null;
            Excel.ListObject? table = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, "Data")
                    ?? throw new InvalidOperationException("Data not found.");
                listObjects = sheet.ListObjects;
                table = listObjects["PlainTable"];
                table.Name = "表1";
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref sheet);
            }
        });

        var info = _tableCommands.Read(batch, "表1");
        RequireSuccess(info);
        Assert.Equal("表1", info.Table!.Name);
        var before = RequireSuccess(_tableCommands.GetData(batch, "表1"));
        Assert.Equal(["Name", "Value"], before.Headers);
        AssertNamedAmountRow(Assert.Single(before.Data), "North", 100);

        var renamed = _tableCommands.Rename(batch, "表1", "テーブル1");
        RequireSuccess(renamed);

        var renamedInfo = _tableCommands.Read(batch, "テーブル1");
        RequireSuccess(renamedInfo);
        Assert.Equal("テーブル1", renamedInfo.Table!.Name);
        Assert.Equal("$A$1:$B$2", renamedInfo.Table.Range);
        var after = RequireSuccess(_tableCommands.GetData(batch, "テーブル1"));
        Assert.Equal(before.Headers, after.Headers);
        Assert.Equal(JsonSerializer.Serialize(before.Data), JsonSerializer.Serialize(after.Data));
        var list = RequireSuccess(_tableCommands.List(batch));
        Assert.DoesNotContain(list.Tables, table => table.Name is "PlainTable" or "表1");
        Assert.Single(list.Tables, table => table.Name == "テーブル1");
        AssertSalesData(RequireSuccess(_tableCommands.GetData(batch, "SalesTable")));
    }

    private static void AssertNumericPair(List<object?> actual, double first, double second)
    {
        Assert.Equal(2, actual.Count);
        Assert.Equal(first, Convert.ToDouble(actual[0], CultureInfo.InvariantCulture));
        Assert.Equal(second, Convert.ToDouble(actual[1], CultureInfo.InvariantCulture));
    }

    private int GetWorksheetTableCount(string sheetName)
    {
        return _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? listObjects = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                listObjects = sheet.ListObjects;
                return listObjects.Count;
            }
            finally
            {
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private object[,] GetRangeValues(string sheetName, string rangeAddress)
    {
        return _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                range = sheet.Range[rangeAddress];
                return (object[,])range.Value2;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
