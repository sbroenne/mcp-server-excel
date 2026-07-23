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

    private static readonly byte[] Ole2Signature =
    [
        0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1
    ];

    private static string? GetConfiguredIrmTestFilePath()
    {
        var irmTestFile = Environment.GetEnvironmentVariable("TEST_IRM_FILE");
        return !string.IsNullOrWhiteSpace(irmTestFile) && System.IO.File.Exists(irmTestFile)
            ? Path.GetFullPath(irmTestFile)
            : null;
    }

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
    public void Test_IrmSignatureFile_DetectsProtection()
    {
        // Arrange - deterministic OLE2 header seam for IRM detection logic
        var testFile = Path.Join(_fixture.TempDir, $"FakeIrm_{Guid.NewGuid():N}.xlsx");
        System.IO.File.WriteAllBytes(testFile, Ole2Signature);

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.True(info.IsValid);
        Assert.True(info.IsIrmProtected);
        Assert.Null(info.Message);
    }

    [Fact]
    public void Test_LegacyXlsFile_NotFlaggedAsIrm()
    {
        // Regression: .xls files are always OLE2 compound documents by design.
        // They must NOT be misclassified as IRM-protected.
        var testFile = Path.Join(_fixture.TempDir, $"LegacyBiff_{Guid.NewGuid():N}.xls");
        System.IO.File.WriteAllBytes(testFile, Ole2Signature);

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        // Note: .xls is not in the Test() valid-extension list (.xlsx/.xlsm only),
        // so IsValid is false. The key assertion is IRM detection.
        Assert.False(info.IsIrmProtected, "Legacy .xls (OLE2) must not be flagged as IRM-protected");
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public void Test_RealIrmProtectedFile_DetectsProtection_WhenConfigured()
    {
        // Real protected workbooks require local credentials and cannot run safely in CI.
        var irmTestFile = GetConfiguredIrmTestFilePath();
        if (irmTestFile == null)
        {
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
