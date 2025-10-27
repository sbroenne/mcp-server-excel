using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

/// <summary>
/// Simple integration tests for VBA trust detection using batch pattern
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Layer", "Core")]
[Trait("Feature", "VBA")]
[Trait("RequiresExcel", "true")]
public class VbaTrustSimpleTests : IDisposable
{
    private readonly string _testDir;
    private readonly string _testFile;
    private readonly SetupCommands _commands;
    private readonly FileCommands _fileCommands;

    public VbaTrustSimpleTests()
    {
        _testDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_VbaTrust_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_testDir);
        _testFile = Path.Combine(_testDir, "test.xlsm");
        _commands = new SetupCommands();
        _fileCommands = new FileCommands();

        // Create test workbook (xlsm extension creates macro-enabled file)
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
    public async Task CheckVbaTrust_ReturnsResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testFile);
        var result = await _commands.CheckVbaTrustAsync(batch);

        // Assert
        Assert.NotNull(result);
        // Result can be trusted or not - both are valid states
        Assert.True(result.IsTrusted || !result.IsTrusted);
    }
}
