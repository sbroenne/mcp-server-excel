using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Setup;

/// <summary>
/// Tests for VBA trust setup operations
/// </summary>
public partial class SetupCommandsTests
{
    [Fact]
    public async Task CheckVbaTrust_ReturnsResult()
    {
        // Arrange - Use .xlsm for macro-enabled file
        var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
            nameof(SetupCommandsTests), nameof(CheckVbaTrust_ReturnsResult), _tempDir, ".xlsm");

        // Act
        await using var batch = await ExcelSession.BeginBatchAsync(testFile);
        var result = await _setupCommands.CheckVbaTrustAsync(batch);

        // Assert
        Assert.NotNull(result);
        // IsTrusted can be true or false depending on system configuration
        // Just verify we got a valid response
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
}
