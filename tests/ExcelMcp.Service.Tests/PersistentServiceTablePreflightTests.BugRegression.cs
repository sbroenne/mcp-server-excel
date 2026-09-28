using System.Globalization;
using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
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
        _rangeCommands.SetValues(
            batch,
            "Data",
            "A1:C2",
            [["Label", "IsActive", "Amount"], ["Initial", true, 1.0]]);
        _tableCommands.Create(
            batch, "Data", "DataTable", "A1:C2", true, "TableStyleLight1");

        var rowsJson =
            """[["NewRow", true, 99.5], ["AnotherRow", false, 0.0]]""";
        var deserializedRows =
            JsonSerializer.Deserialize<List<List<object?>>>(rowsJson)!;

        Assert.IsType<JsonElement>(deserializedRows[0][0]);
        Assert.IsType<JsonElement>(deserializedRows[0][1]);
        Assert.IsType<JsonElement>(deserializedRows[0][2]);

        _tableCommands.Append(batch, "DataTable", deserializedRows);

        var info = _tableCommands.Read(batch, "DataTable");
        Assert.True(info.Success, $"Read after append failed: {info.ErrorMessage}");
        Assert.Equal(3, info.Table!.RowCount);
    }

    [Theory]
    [InlineData("")]
    [InlineData("TableStyleLight1")]
    [InlineData("TableStyleMedium2")]
    public void ListAndRead_WithTableStyle_ReturnsExpectedTableStyle(string tableStyle)
    {
        var batch = _fixture.BatchToken;
        var sheetName = _fixture.CreateTestSheet(batch);
        _rangeCommands.SetValues(
            batch,
            sheetName,
            "A1:B2",
            [["Name", "Value"], ["Example", 1]]);
        _tableCommands.Create(
            batch,
            sheetName,
            "PlainTable",
            "A1:B2");

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

        Assert.True(list.Success, $"List failed: {list.ErrorMessage}");
        Assert.Equal(
            tableStyle,
            Assert.Single(list.Tables, table => table.Name == "PlainTable").TableStyle);
        Assert.True(read.Success, $"Read failed: {read.ErrorMessage}");
        Assert.Equal(tableStyle, read.Table!.TableStyle);
    }

    [Theory]
    [InlineData("表1")]
    [InlineData("テーブル1")]
    [InlineData("표1")]
    public void Create_WithNonAsciiTableName_CreatesReadableTable(string tableName)
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        _rangeCommands.SetValues(
            batch, "Data", "A1:B2", [["Name", "Value"], ["North", 100]]);

        var result = _tableCommands.Create(
            batch, "Data", tableName, "A1:B2", true, "TableStyleLight1");
        Assert.True(result.Success, result.ErrorMessage);

        var info = _tableCommands.Read(batch, tableName);
        Assert.True(info.Success, info.ErrorMessage);
        Assert.NotNull(info.Table);
        Assert.Equal(tableName, info.Table.Name);
        Assert.Equal("Data", info.Table.SheetName);
    }

    [Fact]
    public void Create_WithExcelInvalidTableName_DoesNotLeaveDefaultTable()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        _rangeCommands.SetValues(
            batch, "Data", "A1:B2", [["Name", "Value"], ["North", 100]]);
        var beforeCount = GetWorksheetTableCount("Data");

        Assert.ThrowsAny<Exception>(() =>
            _tableCommands.Create(batch, "Data", new string('A', 256), "A1:B2", true, "TableStyleLight1"));

        Assert.Equal(beforeCount, GetWorksheetTableCount("Data"));
        var values = GetRangeValues("Data", "A1:B2");
        Assert.Equal("Name", values[1, 1]);
        Assert.Equal("Value", values[1, 2]);
        Assert.Equal("North", values[2, 1]);
        Assert.Equal(100d, Convert.ToDouble(values[2, 2], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void Preflight_WithExcelInvalidTableName_ReturnsBlocker()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        _rangeCommands.SetValues(
            batch, "Data", "A1:B2", [["Name", "Value"], ["North", 100]]);

        var result = _tableCommands.Preflight(
            batch, "Data", new string('A', 256), "A1:B2");

        Assert.True(result.Success, result.ErrorMessage);
        Assert.False(result.SafeToCreate);
        var finding = Assert.Single(
            result.Findings,
            item => item.Kind == TablePreflightFindingKind.TableNameInvalid);
        Assert.Equal(TablePreflightSeverity.Blocker, finding.Severity);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Create_RejectedName_PreservesFormulasAndFormatting(bool hasHeaders)
    {
        var batch = _fixture.BatchToken;
        SetUpSourceState(hasHeaders);

        var before = GetSourceState();
        Assert.Throws<ArgumentException>(() =>
            _tableCommands.Create(batch, "Data", new string('A', 256), "A1:B2", hasHeaders));
        var after = GetSourceState();

        Assert.Equal(0, GetWorksheetTableCount("Data"));
        Assert.Equal(before.WorkbookCount, after.WorkbookCount);
        Assert.Equal(before.Formulas.Cast<object>(), after.Formulas.Cast<object>());
        Assert.Equal(before.Formats, after.Formats);
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

    private (object[,] Formulas, string[] Formats, int WorkbookCount) GetSourceState()
    {
        return _fixture.ExecuteRawVerification((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? range = null;
            Microsoft.Office.Interop.Excel.Workbooks? workbooks = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, "Data")
                    ?? throw new InvalidOperationException("Data not found.");
                range = sheet.Range["A1:C4"];
                var formulas = (object[,])range.Formula;
                var formats = new List<string>();
                for (int row = 1; row <= 4; row++)
                {
                    for (int column = 1; column <= 3; column++)
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
        Assert.True(result.Success, result.ErrorMessage);
        var after = GetSourceState();
        Assert.Equal(before.WorkbookCount, after.WorkbookCount);
        Assert.Equal(1, GetWorksheetTableCount("Data"));
        var info = _tableCommands.Read(batch, tableName);
        Assert.True(info.Success, info.ErrorMessage);
        Assert.NotNull(info.Table);
        Assert.Equal(tableName, info.Table.Name);
        Assert.Equal(hasHeaders ? 1 : 2, info.Table.RowCount);
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
        _rangeCommands.SetValues(
            batch, "Data", "A1:B2", [["Name", "Value"], ["North", 100]]);
        var result = _tableCommands.Create(
            batch, "Data", "PlainTable", "A1:B2", true, "TableStyleLight1");
        Assert.True(result.Success, result.ErrorMessage);

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
        Assert.True(info.Success, info.ErrorMessage);
        Assert.Equal("表1", info.Table!.Name);

        var renamed = _tableCommands.Rename(batch, "表1", "テーブル1");
        Assert.True(renamed.Success, renamed.ErrorMessage);

        var renamedInfo = _tableCommands.Read(batch, "テーブル1");
        Assert.True(renamedInfo.Success, renamedInfo.ErrorMessage);
        Assert.Equal("テーブル1", renamedInfo.Table!.Name);
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
