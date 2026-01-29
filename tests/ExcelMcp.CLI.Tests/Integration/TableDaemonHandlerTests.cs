using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Integration tests for Table daemon handlers.
/// Verifies that daemon handlers correctly delegate to Core Commands.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Feature", "Table")]
[Trait("Layer", "CLI")]
public class TableDaemonHandlerTests : DaemonIntegrationTestBase
{
    private readonly TableCommands _tableCommands = new();
    private readonly RangeCommands _rangeCommands = new();
    private readonly SheetCommands _sheetCommands = new();

    public TableDaemonHandlerTests(TempDirectoryFixture fixture) : base(fixture) { }

    [Fact]
    [Trait("Speed", "Fast")]
    public void TableList_EmptyWorkbook_ReturnsEmptyList()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(TableList_EmptyWorkbook_ReturnsEmptyList));
        using var batch = CreateBatch(testFile);

        // Act
        var result = _tableCommands.List(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.Empty(result.Tables);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void TableCreate_ValidRange_CreatesTable()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(TableCreate_ValidRange_CreatesTable));
        using var batch = CreateBatch(testFile);
        var sheets = _sheetCommands.List(batch);
        var sheetName = sheets.Worksheets.First().Name;
        var tableName = $"Table{Guid.NewGuid():N}"[..20];

        // Write header row
        var values = new List<List<object?>>
        {
            new() { "Name", "Value" },
            new() { "Item1", 100 },
            new() { "Item2", 200 }
        };
        _rangeCommands.SetValues(batch, sheetName, "A1:B3", values);

        // Act
        _tableCommands.Create(batch, sheetName, tableName, "A1:B3", hasHeaders: true);

        // Assert
        var listResult = _tableCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.Contains(listResult.Tables, t => t.Name == tableName);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void TableRead_ExistingTable_ReturnsTableInfo()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(TableRead_ExistingTable_ReturnsTableInfo));
        using var batch = CreateBatch(testFile);
        var sheets = _sheetCommands.List(batch);
        var sheetName = sheets.Worksheets.First().Name;
        var tableName = $"Table{Guid.NewGuid():N}"[..20];

        // Create data and table
        var values = new List<List<object?>>
        {
            new() { "Name", "Value" },
            new() { "Item1", 100 }
        };
        _rangeCommands.SetValues(batch, sheetName, "A1:B2", values);
        _tableCommands.Create(batch, sheetName, tableName, "A1:B2", hasHeaders: true);

        // Act
        var result = _tableCommands.Read(batch, tableName);

        // Assert
        Assert.True(result.Success, $"Read failed: {result.ErrorMessage}");
        Assert.NotNull(result.Table);
        Assert.Equal(tableName, result.Table.Name);
        Assert.Equal(2, result.Table.ColumnCount);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void TableRename_ExistingTable_RenamesSuccessfully()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(TableRename_ExistingTable_RenamesSuccessfully));
        using var batch = CreateBatch(testFile);
        var sheets = _sheetCommands.List(batch);
        var sheetName = sheets.Worksheets.First().Name;
        var uniqueId = Guid.NewGuid().ToString("N")[..8];
        var oldName = $"Old{uniqueId}";
        var newName = $"New{uniqueId}";

        // Create data and table
        var values = new List<List<object?>>
        {
            new() { "Name", "Value" },
            new() { "Item1", 100 }
        };
        _rangeCommands.SetValues(batch, sheetName, "A1:B2", values);
        _tableCommands.Create(batch, sheetName, oldName, "A1:B2", hasHeaders: true);

        // Act
        _tableCommands.Rename(batch, oldName, newName);

        // Assert
        var listResult = _tableCommands.List(batch);
        Assert.True(listResult.Success);
        Assert.DoesNotContain(listResult.Tables, t => t.Name == oldName);
        Assert.Contains(listResult.Tables, t => t.Name == newName);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void TableDelete_ExistingTable_DeletesSuccessfully()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(TableDelete_ExistingTable_DeletesSuccessfully));
        using var batch = CreateBatch(testFile);
        var sheets = _sheetCommands.List(batch);
        var sheetName = sheets.Worksheets.First().Name;
        var tableName = $"Table{Guid.NewGuid():N}"[..20];

        // Create data and table
        var values = new List<List<object?>>
        {
            new() { "Name", "Value" },
            new() { "Item1", 100 }
        };
        _rangeCommands.SetValues(batch, sheetName, "A1:B2", values);
        _tableCommands.Create(batch, sheetName, tableName, "A1:B2", hasHeaders: true);

        // Verify table exists
        var beforeList = _tableCommands.List(batch);
        Assert.Contains(beforeList.Tables, t => t.Name == tableName);

        // Act
        _tableCommands.Delete(batch, tableName);

        // Assert
        var afterList = _tableCommands.List(batch);
        Assert.True(afterList.Success);
        Assert.DoesNotContain(afterList.Tables, t => t.Name == tableName);
    }

    [Fact]
    [Trait("Speed", "Fast")]
    public void TableGetData_ExistingTable_ReturnsData()
    {
        // Arrange
        var testFile = CreateTestFile(nameof(TableGetData_ExistingTable_ReturnsData));
        using var batch = CreateBatch(testFile);
        var sheets = _sheetCommands.List(batch);
        var sheetName = sheets.Worksheets.First().Name;
        var tableName = $"Table{Guid.NewGuid():N}"[..20];

        // Create data and table
        var values = new List<List<object?>>
        {
            new() { "Name", "Value" },
            new() { "Item1", 100 },
            new() { "Item2", 200 }
        };
        _rangeCommands.SetValues(batch, sheetName, "A1:B3", values);
        _tableCommands.Create(batch, sheetName, tableName, "A1:B3", hasHeaders: true);

        // Act
        var result = _tableCommands.GetData(batch, tableName, visibleOnly: false);

        // Assert
        Assert.True(result.Success, $"GetData failed: {result.ErrorMessage}");
        Assert.Equal(2, result.RowCount); // 2 data rows (excluding header)
    }
}
