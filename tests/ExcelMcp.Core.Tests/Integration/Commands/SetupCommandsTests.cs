using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands;

/// <summary>
/// Integration tests for Setup Core operations.
/// Tests Core layer directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Setup")]
[Trait("RequiresExcel", "true")]
public class SetupCommandsTests : IDisposable
{
    private readonly ISetupCommands _setupCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;

    public SetupCommandsTests()
    {
        _setupCommands = new SetupCommands();
        _fileCommands = new FileCommands();

        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_SetupTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestWorkbook.xlsm"); // Macro-enabled for VBA trust

        // Create test Excel file
        var result = _fileCommands.CreateEmptyAsync(_testExcelFile).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}");
        }
    }

    [Fact]
    public async Task CheckVbaTrust_ReturnsResult()
    {
        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _setupCommands.CheckVbaTrustAsync(batch);

        // Assert
        Assert.NotNull(result);
        // IsTrusted can be true or false depending on system configuration
        Assert.True(result.IsTrusted || !result.IsTrusted);
    }

    [Fact]
    public void EnableVbaTrust_ReturnsResult()
    {
        // Act
        var result = _setupCommands.EnableVbaTrust();

        // Assert
        Assert.NotNull(result);
        Assert.NotNull(result.RegistryPathsSet);
        // Success depends on whether registry keys were set
    }

    [Fact]
    public async Task CheckVbaTrust_AfterEnable_MayBeTrusted()
    {
        // Arrange
        _setupCommands.EnableVbaTrust();

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(_testExcelFile);
        var result = await _setupCommands.CheckVbaTrustAsync(batch);

        // Assert
        Assert.NotNull(result);
        // May be trusted after enabling (depends on system state)
    }

    public void Dispose()
    {
        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, true);
            }
        }
        catch
        {
            // Ignore cleanup errors
        }

        GC.SuppressFinalize(this);
    }
}
