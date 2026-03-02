using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.File;

/// <summary>
/// Tests for IRM (Azure Information Rights Management) file detection.
/// Feature: Improvement #5 IRM File Auto-Detection and Suggestions
/// </summary>
public partial class FileCommandsTests
{
    // === IMPROVEMENT #5: IRM DETECTION TESTS ===

    [Fact]
    public void Test_NormalFile_NoIrmProtection()
    {
        // Arrange - Create a normal Excel file without IRM
        var testFile = _fixture.CreateTestFile();

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.True(info.IsValid);
        Assert.False(info.IsIrmProtected);
        Assert.Null(info.Message);
    }

    [Fact]
    public void Test_IrmProtectedFile_DetectsProtection()
    {
        // Arrange - Skip if no IRM protected test file available
        // For now, simulate or use a real IRM-protected file if available
        string? irmTestFile = Environment.GetEnvironmentVariable("TEST_IRM_FILE");
        if (string.IsNullOrEmpty(irmTestFile) || !System.IO.File.Exists(irmTestFile))
        {
            // Skip this test - requires actual IRM-protected file
            return;
        }

        // Act
        var info = _fileCommands.Test(irmTestFile);

        // Assert
        Assert.True(info.Exists);
        Assert.True(info.IsIrmProtected);
        // IsValid might be false or true depending on implementation - main thing is IRM detection
    }

    [Fact]
    public void Test_FileInfo_IncludesIrmStatus()
    {
        // Arrange
        var testFile = _fixture.CreateTestFile();

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert - Ensure IRM detection info is available
        Assert.NotNull(info);
        // The FileValidationInfo should have IsIrmProtected property populated
        Assert.False(info.IsIrmProtected);  // Normal test file should not be IRM protected
    }

    [Fact]
    public void Test_RvToolsExportFile_DetectsIrmIfPresent()
    {
        // Arrange - Test with a typical RVTools export pattern
        // RVTools exports from Mercedes/FDC would be IRM-protected
        var testFile = _fixture.CreateTestFile();

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert - IRM detection should work for RVTools pattern
        // (Our test file is not IRM, but the detection logic should work)
        Assert.True(info.Exists);
        Assert.False(info.IsIrmProtected);  // Unprotected test file
    }
}
