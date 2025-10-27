using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Simple integration tests for CellCommands using batch pattern
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Cells")]
[Trait("RequiresExcel", "true")]
public class CellCommandsSimpleTests : IDisposable
{
    private readonly string _testDir;
    private readonly string _testFile;
    private readonly CellCommands _commands;
    private readonly FileCommands _fileCommands;

    public CellCommandsSimpleTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_CellSimple_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
        _testFile = Path.Combine(_testDir, "test.xlsx");
        _commands = new CellCommands();
        _fileCommands = new FileCommands();

        // Create test workbook
        var result = _fileCommands.CreateEmptyAsync(_testFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test file: {result.ErrorMessage}");
        }
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_testDir))
            {
                Directory.Delete(_testDir, recursive: true);
            }
        }
        catch { /* Cleanup failure is non-critical */ }
        GC.SuppressFinalize(this);
    }

    [Fact]
    public async Task GetValue_EmptyCell_ReturnsNull()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.GetValueAsync(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success, $"GetValue failed: {result.ErrorMessage}");
        Assert.Null(result.Value);
    }

    [Fact]
    public async Task SetValue_ValidCell_Success()
    {
        // Arrange
        const string testValue = "Test Value";

        // Act - Set value
        await using (var batch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var setResult = await _commands.SetValueAsync(batch, "Sheet1", "A1", testValue);
            Assert.True(setResult.Success, $"SetValue failed: {setResult.ErrorMessage}");
            await batch.SaveAsync();
        }

        // Act - Get value (new batch)
        await using (var batch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var getResult = await _commands.GetValueAsync(batch, "Sheet1", "A1");

            // Assert
            Assert.True(getResult.Success, $"GetValue failed: {getResult.ErrorMessage}");
            Assert.Equal(testValue, getResult.Value?.ToString());
        }
    }

    [Fact]
    public async Task GetFormula_EmptyCell_ReturnsNull()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.GetFormulaAsync(batch, "Sheet1", "A1");

        // Assert
        Assert.True(result.Success, $"GetFormula failed: {result.ErrorMessage}");
        Assert.Null(result.Formula);
    }
}
