using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;
using ExcelListObjects = Microsoft.Office.Interop.Excel.ListObjects;
using ExcelRange = Microsoft.Office.Interop.Excel.Range;
using ExcelWorksheet = Microsoft.Office.Interop.Excel.Worksheet;

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
    /// Regression test for issue #894:
    /// Excel accepts localized non-ASCII table names such as 表1, so create must not
    /// reject them before Excel can create the table.
    /// </summary>
    [Fact]
    public void Create_WithNonAsciiTableName_CreatesReadableTable()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(TableCommandsTests_BugRegression),
            nameof(Create_WithNonAsciiTableName_CreatesReadableTable),
            _fixture.TempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            ExcelWorksheet? sheet = null;
            ExcelRange? dataRange = null;
            try
            {
                sheet = (ExcelWorksheet)ctx.Book.Worksheets[1];
                sheet.Name = "Data";
                dataRange = sheet.Range["A1:B2"];
                dataRange.Value2 = new object[,]
                {
                    { "Name", "Value" },
                    { "North", 100 },
                };
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref dataRange);
                ComUtilities.Release(ref sheet);
            }
        });

        _tableCommands.Create(batch, "Data", "表1", "A1:B2", true, "TableStyleLight1");

        var info = _tableCommands.Read(batch, "表1");
        Assert.True(info.Success, info.ErrorMessage);
        Assert.NotNull(info.Table);
        Assert.Equal("表1", info.Table.Name);
        Assert.Equal("Data", info.Table.SheetName);
    }

    /// <summary>
    /// Regression test for failed Excel name assignment leaving behind the default
    /// table created before Excel rejected the requested name.
    /// </summary>
    [Fact]
    public void Create_WithExcelInvalidTableName_DoesNotLeaveDefaultTable()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(TableCommandsTests_BugRegression),
            nameof(Create_WithExcelInvalidTableName_DoesNotLeaveDefaultTable),
            _fixture.TempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            ExcelWorksheet? sheet = null;
            ExcelRange? dataRange = null;
            try
            {
                sheet = (ExcelWorksheet)ctx.Book.Worksheets[1];
                sheet.Name = "Data";
                dataRange = sheet.Range["A1:B2"];
                dataRange.Value2 = new object[,]
                {
                    { "Name", "Value" },
                    { "North", 100 },
                };
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref dataRange);
                ComUtilities.Release(ref sheet);
            }
        });

        var beforeCount = GetWorksheetTableCount(batch, "Data");

        Assert.ThrowsAny<Exception>(() =>
            _tableCommands.Create(batch, "Data", "Invalid Name", "A1:B2", true, "TableStyleLight1"));

        Assert.Equal(beforeCount, GetWorksheetTableCount(batch, "Data"));
    }

    /// <summary>
    /// Regression test for issue #894:
    /// Existing Excel-created localized table names must be readable and renamable
    /// through the shared table commands.
    /// </summary>
    [Fact]
    public void ReadAndRename_WithExistingNonAsciiTableName_Succeeds()
    {
        var testFile = CoreTestHelper.CreateUniqueTestFile(
            nameof(TableCommandsTests_BugRegression),
            nameof(ReadAndRename_WithExistingNonAsciiTableName_Succeeds),
            _fixture.TempDir,
            ".xlsx");

        using var batch = ExcelSession.BeginBatch(testFile);

        batch.Execute((ctx, ct) =>
        {
            ExcelWorksheet? sheet = null;
            ExcelRange? dataRange = null;
            try
            {
                sheet = (ExcelWorksheet)ctx.Book.Worksheets[1];
                sheet.Name = "Data";
                dataRange = sheet.Range["A1:B2"];
                dataRange.Value2 = new object[,]
                {
                    { "Name", "Value" },
                    { "North", 100 },
                };
                return 0;
            }
            finally
            {
                ComUtilities.Release(ref dataRange);
                ComUtilities.Release(ref sheet);
            }
        });
        _tableCommands.Create(batch, "Data", "PlainTable", "A1:B2", true, "TableStyleLight1");

        batch.Execute((ctx, ct) =>
        {
            dynamic? sheet = null;
            dynamic? listObjects = null;
            dynamic? table = null;
            try
            {
                sheet = ctx.Book.Worksheets[1];
                listObjects = sheet.ListObjects;
                table = listObjects.Item("PlainTable");
                table.Name = "表1";
                return 0;
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

        _tableCommands.Rename(batch, "表1", "テーブル1");

        var renamedInfo = _tableCommands.Read(batch, "テーブル1");
        Assert.True(renamedInfo.Success, renamedInfo.ErrorMessage);
        Assert.Equal("テーブル1", renamedInfo.Table!.Name);
    }

    private static int GetWorksheetTableCount(IExcelBatch batch, string sheetName)
    {
        return batch.Execute((ctx, ct) =>
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
