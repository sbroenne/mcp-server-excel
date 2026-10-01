using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.Service.Mac;
using Xunit;

namespace Sbroenne.ExcelMcp.Portable.Tests;

public sealed class MacAutomationProbeTests
{
    [Theory]
    [InlineData(0, "Allowed")]
    [InlineData(-1743, "Denied")]
    [InlineData(-1744, "ConsentRequired")]
    [InlineData(-600, "ExcelNotRunning")]
    [InlineData(-1712, "Error")]
    [InlineData(123, "Error")]
    public void PermissionStatus_MapsNativeResultWithoutSuccessFallback(int code, string expected)
    {
        Assert.Equal(expected, MacAutomationAccess.DescribeStatus(code));
    }

    [Fact]
    public void AppleEventDescriptor_MatchesPackedMacSdkLayout()
    {
        Assert.Equal(4 + IntPtr.Size, Marshal.SizeOf<MacAutomationAccess.Descriptor>());
        Assert.Equal(4, Marshal.OffsetOf<MacAutomationAccess.Descriptor>(
            nameof(MacAutomationAccess.Descriptor.DataHandle)).ToInt32());
    }
}
