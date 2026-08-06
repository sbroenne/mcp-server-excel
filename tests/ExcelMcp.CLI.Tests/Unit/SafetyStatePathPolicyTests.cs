using Sbroenne.ExcelMcp.Service.Safety;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Unit;

[Trait("Layer", "Service")]
[Trait("Category", "Unit")]
[Trait("Feature", "SafetyPathPolicy")]
[Trait("RequiresExcel", "false")]
[Trait("Speed", "Fast")]
public sealed class SafetyStatePathPolicyTests : IDisposable
{
    private readonly string _stateRoot = Path.Combine(
        Path.GetTempPath(),
        $"excelmcp-safety-path-{Guid.NewGuid():N}");

    [Fact]
    public void PrepareRoot_AcceptsOrdinaryLocalDirectory()
    {
        var result = SafetyStatePathPolicy.PrepareRoot(_stateRoot);

        Assert.Equal(Path.GetFullPath(_stateRoot), result);
        Assert.True(Directory.Exists(result));
    }

    [Fact]
    public void PrepareRoot_RejectsNetworkPathBeforeAccessingIt()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            SafetyStatePathPolicy.PrepareRoot(@"\\server.invalid\shared\excelmcp"));

        Assert.Contains("local", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PrepareRoot_ValidatesExistingAncestorsBeforeCreatingDirectory()
    {
        var created = false;

        Assert.Throws<InvalidOperationException>(() => SafetyStatePathPolicy.PrepareRoot(
            _stateRoot,
            _ => throw new InvalidOperationException("synthetic reparse-point rejection"),
            _ => created = true));

        Assert.False(created);
        Assert.False(Directory.Exists(_stateRoot));
    }

    [Fact]
    public void IsNetworkDriveType_RejectsMappedNetworkDrives()
    {
        Assert.True(SafetyStatePathPolicy.IsNetworkDriveType(DriveType.Network));
        Assert.False(SafetyStatePathPolicy.IsNetworkDriveType(DriveType.Fixed));
    }

    [Theory]
    [InlineData(FileAttributes.Directory, false)]
    [InlineData(FileAttributes.Directory | FileAttributes.ReparsePoint, true)]
    [InlineData(FileAttributes.ReparsePoint, true)]
    public void IsReparsePoint_ClassifiesAttributes(FileAttributes attributes, bool expected)
    {
        Assert.Equal(expected, SafetyStatePathPolicy.IsReparsePoint(attributes));
    }

    [Fact]
    public void TryGetExistingAttributes_DistinguishesMissingPathFromExistingEntry()
    {
        Directory.CreateDirectory(_stateRoot);

        Assert.True(SafetyStatePathPolicy.TryGetExistingAttributes(_stateRoot, out var attributes));
        Assert.True((attributes & FileAttributes.Directory) != 0);
        Assert.False(SafetyStatePathPolicy.TryGetExistingAttributes(
            Path.Combine(_stateRoot, "missing"),
            out _));
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (Directory.Exists(_stateRoot))
        {
            Directory.Delete(_stateRoot, recursive: true);
        }

        GC.SuppressFinalize(this);
    }
}
