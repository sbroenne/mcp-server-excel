using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Simple integration tests for SheetCommands using batch pattern
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Worksheets")]
[Trait("RequiresExcel", "true")]
public class SheetCommandsSimpleTests : IDisposable
{
    private readonly string _testDir;
    private readonly string _testFile;
    private readonly SheetCommands _commands;
    private readonly FileCommands _fileCommands;

    public SheetCommandsSimpleTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_SheetSimple_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
        _testFile = Path.Combine(_testDir, "test.xlsx");
        _commands = new SheetCommands();
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
    public async Task List_NewWorkbook_ReturnsDefaultSheet()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Worksheets);
        Assert.NotEmpty(result.Worksheets);
        Assert.Contains(result.Worksheets, s => s.Name == "Sheet1");
    }

    [Fact]
    public async Task Create_NewSheet_Success()
    {
        // Arrange
        const string newSheetName = "TestSheet";

        // Act - Create sheet
        await using (var batch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var createResult = await _commands.CreateAsync(batch, newSheetName);
            Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
            await batch.SaveAsync();
        }

        // Act - List sheets (new batch)
        await using (var batch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var listResult = await _commands.ListAsync(batch);

            // Assert
            Assert.True(listResult.Success, $"List failed: {listResult.ErrorMessage}");
            Assert.Contains(listResult.Worksheets!, s => s.Name == newSheetName);
        }
    }

    [Fact]
    public async Task Read_EmptySheet_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ReadAsync(batch, "Sheet1", "A1:B2");

        // Assert
        Assert.True(result.Success, $"Read failed: {result.ErrorMessage}");
        Assert.NotNull(result.Data);
    }
}
