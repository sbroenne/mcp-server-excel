using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Table;

/// <summary>
/// Bug regression tests for TableCommands.
/// These tests reproduce known bugs and must fail before the fix and pass after.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Tables")]
[Trait("Speed", "Medium")]
public sealed class TableCommandsTests_BugRegression : IClassFixture<TempDirectoryFixture>
{
    private readonly TableCommands _tableCommands;
    private readonly TempDirectoryFixture _fixture;

    /// <summary>
    /// Initializes a new instance of the <see cref="TableCommandsTests_BugRegression"/> class.
    /// </summary>
    public TableCommandsTests_BugRegression(TempDirectoryFixture fixture)
    {
        _tableCommands = new TableCommands();
        _fixture = fixture;
    }

    /// <summary>
    /// Regression test for issue #519:
    /// table append throws COM marshalling exception when row values are JsonElement
    /// (as produced by CLI JSON deserialization of --rows parameter).
    /// Before fix: throws NotSupportedException / InvalidCastException / COMException.
    /// After fix: appends rows successfully.
    /// </summary>
    [Fact]
    public void Append_WithJsonElementValues_DoesNotThrow()
    {
        // Arrange: create a workbook with a table that has string, bool, and number columns
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(TableCommandsTests_BugRegression),
            nameof(Append_WithJsonElementValues_DoesNotThrow),
            _fixture.TempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        // Create data + table in the same batch (no save needed)
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            sheet.Name = "Data";
            sheet.Range["A1"].Value2 = "Label";
            sheet.Range["B1"].Value2 = "IsActive";
            sheet.Range["C1"].Value2 = "Amount";
            sheet.Range["A2"].Value2 = "Initial";
            sheet.Range["B2"].Value2 = true;
            sheet.Range["C2"].Value2 = 1.0;
            return 0;
        });
        _tableCommands.Create(batch, "Data", "DataTable", "A1:C2", true, "TableStyleLight1");

        // Act: deserialize rows the same way the CLI does — via JsonSerializer producing JsonElement
        // This is key: the values must be JsonElement (boxed as object?), not raw C# types
        var rowsJson = """[["NewRow", true, 99.5], ["AnotherRow", false, 0.0]]""";
        var deserializedRows = JsonSerializer.Deserialize<List<List<object?>>>(rowsJson)!;

        // Confirm the test is correctly structured: values must be JsonElements, not strings/bools
        Assert.IsType<JsonElement>(deserializedRows[0][0]);
        Assert.IsType<JsonElement>(deserializedRows[0][1]);
        Assert.IsType<JsonElement>(deserializedRows[0][2]);

        // Assert: should not throw — before the fix this throws a COM marshalling exception
        _tableCommands.Append(batch, "DataTable", deserializedRows);

        // Verify rows were appended
        var info = _tableCommands.Read(batch, "DataTable");
        Assert.True(info.Success, $"Read after append failed: {info.ErrorMessage}");
        Assert.Equal(3, info.Table!.RowCount); // 1 original + 2 appended
    }

    /// <summary>
    /// Regression test for issue #893: table list and read support tables without a style.
    /// </summary>
    [Fact]
    public void ListAndRead_WithUnstyledTable_ReturnsEmptyTableStyle()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(TableCommandsTests_BugRegression),
            nameof(ListAndRead_WithUnstyledTable_ReturnsEmptyTableStyle),
            _fixture.TempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);
        batch.Execute((ctx, ct) =>
        {
            dynamic? worksheets = null;
            dynamic? sheet = null;
            dynamic? range = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
                sheet = worksheets[1];
                range = sheet.Range["A1:B2"];
                range.Value2 = new object[,] { { "Name", "Value" }, { "Example", 1 } };
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref range);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref worksheets);
            }
        });
        _tableCommands.Create(batch, "Sheet1", "PlainTable", "A1:B2");
        batch.Execute((ctx, ct) =>
        {
            dynamic? worksheets = null;
            dynamic? sheet = null;
            dynamic? listObjects = null;
            dynamic? table = null;
            try
            {
                worksheets = ctx.Book.Worksheets;
                sheet = worksheets[1];
                listObjects = sheet.ListObjects;
                table = listObjects["PlainTable"];
                table.TableStyle = "";
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref worksheets);
            }
        });

        var list = _tableCommands.List(batch);
        var read = _tableCommands.Read(batch, "PlainTable");

        Assert.True(list.Success, $"List failed: {list.ErrorMessage}");
        Assert.Equal("", Assert.Single(list.Tables).TableStyle);
        Assert.True(read.Success, $"Read failed: {read.ErrorMessage}");
        Assert.Equal("", read.Table!.TableStyle);
    }
}
