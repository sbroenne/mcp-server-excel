using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Tests for TOM calculated column CRUD operations (LLM-essential only)
/// </summary>
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
public class DataModelTomCommandsTests : IDisposable
{
    private readonly string _tempDir;
    private readonly IDataModelTomCommands _tomCommands;
    private readonly string _testFile;

    public DataModelTomCommandsTests(ITestOutputHelper output)
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"TomTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
        _tomCommands = new DataModelTomCommands();
        _testFile = Path.Combine(_tempDir, "TomTest.xlsx");
        CreateTestFile().GetAwaiter().GetResult();
    }

    private async Task CreateTestFile()
    {
        var fileCommands = new FileCommands();
        var createResult = await fileCommands.CreateEmptyAsync(_testFile);
        if (!createResult.Success)
            throw new InvalidOperationException($"Failed to create test file: {createResult.ErrorMessage}");

        var tableCommands = new TableCommands();

        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        await tableCommands.CreateAsync(batch, "Sheet1", "TestTable", "A1:C4", hasHeaders: true);
        await tableCommands.AddToDataModelAsync(batch, "TestTable");
        await batch.SaveAsync();
    }

    [Fact]
    public void CreateCalculatedColumn_CreatesSuccessfully()
    {
        var result = _tomCommands.CreateCalculatedColumn(
            _testFile, "TestTable", "DoubleAmount", "[Amount] * 2", dataType: "Double");
        Assert.True(result.Success, result.ErrorMessage);
    }

    [Fact]
    public void ListCalculatedColumns_ReturnsColumns()
    {
        _tomCommands.CreateCalculatedColumn(_testFile, "TestTable", "TestCol", "[Amount] + 10");
        var result = _tomCommands.ListCalculatedColumns(_testFile, "TestTable");
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains(result.CalculatedColumns, c => c.Name == "TestCol");
    }

    [Fact]
    public void ViewCalculatedColumn_ReturnsDetails()
    {
        _tomCommands.CreateCalculatedColumn(_testFile, "TestTable", "TripleAmount", "[Amount] * 3");
        var result = _tomCommands.ViewCalculatedColumn(_testFile, "TestTable", "TripleAmount");
        Assert.True(result.Success, result.ErrorMessage);
        Assert.Contains("[Amount]", result.DaxFormula);
    }

    [Fact]
    public void UpdateCalculatedColumn_UpdatesSuccessfully()
    {
        _tomCommands.CreateCalculatedColumn(_testFile, "TestTable", "UpdateTest", "[Amount] + 5");
        var result = _tomCommands.UpdateCalculatedColumn(_testFile, "TestTable", "UpdateTest", daxFormula: "[Amount] * 10");
        Assert.True(result.Success, result.ErrorMessage);
    }

    [Fact]
    public void DeleteCalculatedColumn_DeletesSuccessfully()
    {
        _tomCommands.CreateCalculatedColumn(_testFile, "TestTable", "DeleteTest", "[Amount] - 50");
        var result = _tomCommands.DeleteCalculatedColumn(_testFile, "TestTable", "DeleteTest");
        Assert.True(result.Success, result.ErrorMessage);
    }

    public void Dispose()
    {
        try { if (Directory.Exists(_tempDir)) Directory.Delete(_tempDir, true); } catch { }
        GC.SuppressFinalize(this);
    }
}
