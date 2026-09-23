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

    [Fact]
    public void ListAndRead_WithUnstyledTable_ReturnsEmptyTableStyle()
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
            Excel.Worksheet? sheet = null;
            Excel.ListObjects? tables = null;
            Excel.ListObject? table = null;
            try
            {
                sheet = (Excel.Worksheet)ctx.Book.Worksheets[sheetName];
                tables = sheet.ListObjects;
                table = tables["PlainTable"];
                table.TableStyle = "";
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref tables);
                ComUtilities.Release(ref sheet);
            }
        });

        var list = _tableCommands.List(batch);
        var read = _tableCommands.Read(batch, "PlainTable");

        Assert.True(list.Success, $"List failed: {list.ErrorMessage}");
        Assert.Equal(
            "",
            Assert.Single(list.Tables, table => table.Name == "PlainTable").TableStyle);
        Assert.True(read.Success, $"Read failed: {read.ErrorMessage}");
        Assert.Equal("", read.Table!.TableStyle);
    }
}
