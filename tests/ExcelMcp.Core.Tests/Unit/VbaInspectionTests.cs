using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Unit;

[Trait("Layer", "Core")]
[Trait("Category", "Unit")]
[Trait("Feature", "VBA")]
[Trait("Speed", "Fast")]
[Trait("RequiresExcel", "false")]
public sealed class VbaInspectionTests
{
    [Fact]
    public void ReadReferenceInfo_BrokenReference_DoesNotAccessUnavailableMetadata()
    {
        var info = VbaCommands.ReadReferenceInfo(new BrokenReference(), 4);
        Assert.Equal(4, info.Index);
        Assert.True(info.IsBroken);
        Assert.Null(info.Name);
        Assert.Null(info.Description);
        Assert.Null(info.LibraryId);
        Assert.Null(info.Major);
        Assert.Null(info.Minor);
        Assert.Null(info.BuiltIn);
    }

    [Theory]
    [InlineData(0, "Run")]
    [InlineData(1, "Break")]
    [InlineData(2, "Design")]
    public void GetProjectMode_MapsNativeValues(int mode, string expected)
    {
        Assert.Equal(expected, VbaCommands.GetProjectMode(mode));
    }

    [Theory]
    [InlineData(0, "None")]
    [InlineData(1, "Locked")]
    public void GetProjectProtection_MapsNativeValues(int protection, string expected)
    {
        Assert.Equal(expected, VbaCommands.GetProjectProtection(protection));
    }

    [Fact]
    public void ProjectState_UnknownNativeValues_AreNotReportedAsHealthy()
    {
        Assert.Throws<InvalidOperationException>(() => VbaCommands.GetProjectMode(99));
        Assert.Throws<InvalidOperationException>(() => VbaCommands.GetProjectProtection(99));
    }

    public sealed class BrokenReference
    {
        private readonly InvalidOperationException _metadataError = new("Broken metadata cannot be read.");
        public bool IsBroken { get; } = true;
        public string Name => throw _metadataError;
        public string Description => throw _metadataError;
        public int Major => throw _metadataError;
        public int Minor => throw _metadataError;
        public bool BuiltIn => throw _metadataError;
    }
}
