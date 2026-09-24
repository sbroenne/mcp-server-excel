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

    [Fact]
    public void Create_WithNonAsciiTableName_CreatesReadableTable()
    {
        var batch = _fixture.BatchToken;
        _fixture.CreateNamedTestSheet(batch, "Data");
        _rangeCommands.SetValues(
            batch, "Data", "A1:B2", [["Name", "Value"], ["North", 100]]);

        var result = _tableCommands.Create(
            batch, "Data", "表1", "A1:B2", true, "TableStyleLight1");
        Assert.True(result.Success, result.ErrorMessage);

        var info = _tableCommands.Read(batch, "表1");
        Assert.True(info.Success, info.ErrorMessage);
        Assert.NotNull(info.Table);
        Assert.Equal("表1", info.Table.Name);
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
            _tableCommands.Create(batch, "Data", "Invalid Name", "A1:B2", true, "TableStyleLight1"));

        Assert.Equal(beforeCount, GetWorksheetTableCount("Data"));
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
            ExcelWorksheet? sheet = null;
            ExcelListObjects? listObjects = null;
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
}
