using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Simple integration tests for ParameterCommands using batch pattern
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public class ParameterCommandsSimpleTests : IDisposable
{
    private readonly string _testDir;
    private readonly string _testFile;
    private readonly ParameterCommands _commands;
    private readonly FileCommands _fileCommands;

    public ParameterCommandsSimpleTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_ParamSimple_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
        _testFile = Path.Combine(_testDir, "test.xlsx");
        _commands = new ParameterCommands();
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
    public async Task List_EmptyWorkbook_ReturnsSuccess()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.ListAsync(batch);

        // Assert
        Assert.True(result.Success, $"List failed: {result.ErrorMessage}");
        Assert.NotNull(result.Parameters);
        Assert.Empty(result.Parameters);
    }

    [Fact]
    public async Task Create_NewParameter_Success()
    {
        // Arrange
        const string paramName = "TestParam";
        const string testValue = "TestValue123";

        // Arrange - Put a value in A1 first
        await using (var batch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            await batch.ExecuteAsync<int>((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                sheet.Range["A1"].Value2 = testValue;
                return ValueTask.FromResult(0);
            });
            await batch.SaveAsync();
        }

        // Act - Create parameter
        await using (var batch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var createResult = await _commands.CreateAsync(batch, paramName, "Sheet1!A1");
            Assert.True(createResult.Success, $"Create failed: {createResult.ErrorMessage}");
            await batch.SaveAsync();
        }

        // Act - Get parameter (new batch)
        await using (var batch = await ExcelSession.BeginBatchAsync(_testFile))
        {
            var getResult = await _commands.GetAsync(batch, paramName);

            // Assert
            Assert.True(getResult.Success, $"Get failed: {getResult.ErrorMessage}");
            Assert.NotNull(getResult.Value);
            Assert.Equal(testValue, getResult.Value?.ToString());
        }
    }

    [Fact]
    public async Task Get_NonExistentParameter_ReturnsError()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.GetAsync(batch, "NonExistent");

        // Assert
        Assert.False(result.Success);
        Assert.NotNull(result.ErrorMessage);
    }
}
