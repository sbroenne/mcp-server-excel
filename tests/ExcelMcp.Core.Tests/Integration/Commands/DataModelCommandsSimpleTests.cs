using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Simple integration tests for DataModelCommands using batch pattern
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "DataModel")]
[Trait("RequiresExcel", "true")]
public class DataModelCommandsSimpleTests : IDisposable
{
    private readonly string _testDir;
    private readonly string _testFile;
    private readonly DataModelCommands _commands;
    private readonly FileCommands _fileCommands;

    public DataModelCommandsSimpleTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_DMSimple_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
        _testFile = Path.Combine(_testDir, "test.xlsx");
        _commands = new DataModelCommands();
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
    public async Task ListTables_EmptyDataModel_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListTablesAsync(batch);

        // Assert
        Assert.True(result.Success, $"ListTables failed: {result.ErrorMessage}");
        Assert.NotNull(result.Tables);
        Assert.Empty(result.Tables);
    }

    [Fact]
    public async Task ListMeasures_EmptyDataModel_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListMeasuresAsync(batch);

        // Assert
        Assert.True(result.Success, $"ListMeasures failed: {result.ErrorMessage}");
        Assert.NotNull(result.Measures);
        Assert.Empty(result.Measures);
    }

    [Fact]
    public async Task ListRelationships_EmptyDataModel_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListRelationshipsAsync(batch);

        // Assert
        Assert.True(result.Success, $"ListRelationships failed: {result.ErrorMessage}");
        Assert.NotNull(result.Relationships);
        Assert.Empty(result.Relationships);
    }
}
