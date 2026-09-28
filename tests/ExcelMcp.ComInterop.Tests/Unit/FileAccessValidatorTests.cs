using Sbroenne.ExcelMcp.ComInterop.Tests.Helpers;
using Sbroenne.ExcelMcp.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Unit;

[Trait("Layer", "ComInterop")]
[Trait("Category", "Unit")]
[Trait("Feature", "File")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class FileAccessValidatorTests(
    TempDirectoryFixture fixture) : IClassFixture<TempDirectoryFixture>
{
    [Fact]
    public void IsIrmProtected_PasswordEncryptionMarkersWithoutDrmDataSpace_ReturnsFalse()
    {
        var filePath = fixture.CreateFilePath();
        var bytes = new byte[2048];
        new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }
            .CopyTo(bytes, 0);
        System.Text.Encoding.Unicode.GetBytes("EncryptionInfo")
            .CopyTo(bytes, 512);
        System.Text.Encoding.Unicode.GetBytes("EncryptedPackage")
            .CopyTo(bytes, 1024);
        File.WriteAllBytes(filePath, bytes);

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_LegacyDrmDataSpaceMetadata_ReturnsTrue()
    {
        var filePath = OleDataSpaceTestFile.Write(
            fixture.CreateFilePath(),
            "\tDRMDataSpace");

        Assert.True(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_ModernDrmEncryptedDataSpaceMetadata_ReturnsTrue()
    {
        var filePath = OleDataSpaceTestFile.Write(
            fixture.CreateFilePath(),
            "DRMEncryptedDataSpace");

        Assert.True(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_SimilarDataSpaceName_ReturnsFalse()
    {
        var filePath = OleDataSpaceTestFile.Write(
            fixture.CreateFilePath(),
            "DRMEncryptedDataSpacePreview");

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_MapWithoutMatchingDefinition_ReturnsFalse()
    {
        var filePath = OleDataSpaceTestFile.Write(
            fixture.CreateFilePath(),
            "\tDRMDataSpace",
            "DifferentDataSpace");

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_DeepDirectoryTree_ReturnsFalseWithoutRecursionFailure()
    {
        var filePath = OleDataSpaceTestFile.WriteDeepDirectory(
            fixture.CreateFilePath(),
            entryCount: 10_000);

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void IsIrmProtected_ThousandsOfNestedMaximumLengthStorages_HasBoundedAllocation()
    {
        var filePath = OleDataSpaceTestFile.WriteDeepNestedStorages(
            fixture.CreateFilePath(),
            entryCount: 2_000);
        var before = GC.GetAllocatedBytesForCurrentThread();

        var isIrmProtected = FileAccessValidator.IsIrmProtected(filePath);

        var allocatedBytes = GC.GetAllocatedBytesForCurrentThread() - before;
        Assert.False(isIrmProtected);
        Assert.True(
            allocatedBytes < 32 * 1024 * 1024,
            $"Nested OLE inspection allocated {allocatedBytes:N0} bytes.");
    }

    [Fact]
    public void IsIrmProtected_OversizedVersion4RootMiniStream_ReturnsFalse()
    {
        var filePath = OleDataSpaceTestFile.WriteOversizedVersion4RootMiniStream(
            fixture.CreateFilePath());

        Assert.False(FileAccessValidator.IsIrmProtected(filePath));
    }

    [Fact]
    public void GetSectorCount_SparseFileBeyondInspectionLimit_ThrowsInvalidDataException()
    {
        var oversizedLength = checked(((long)int.MaxValue + 2) * 512);

        var exception = Assert.Throws<InvalidDataException>(
            () => OleCompoundFileReader.GetSectorCount(oversizedLength, 512));

        Assert.Contains("length", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

}
