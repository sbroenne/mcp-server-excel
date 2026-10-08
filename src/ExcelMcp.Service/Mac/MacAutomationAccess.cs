using System.Runtime.InteropServices;
using System.Text;

namespace Sbroenne.ExcelMcp.Service.Mac;

public static class MacAutomationAccess
{
    private const string AppleEvents =
        "/System/Library/Frameworks/CoreServices.framework/Frameworks/AE.framework/AE";

    // AEDataModel.h uses two-byte packing, including on 64-bit macOS.
    [StructLayout(LayoutKind.Sequential, Pack = 2)]
    public struct Descriptor
    {
        public uint DescriptorType;
        public System.IntPtr DataHandle;
    }

    public static string DescribeStatus(int status)
    {
        return status switch
        {
            0 => "Allowed",
            -1743 => "Denied",
            -1744 => "ConsentRequired",
            -600 => "ExcelNotRunning",
            _ => "Error"
        };
    }

    public static int Check()
    {
        if (!System.OperatingSystem.IsMacOS())
        {
            throw new System.PlatformNotSupportedException("Apple Events require macOS.");
        }

        var bundleId = Encoding.UTF8.GetBytes("com.microsoft.Excel");
        int status = AECreateDesc(0x62756E64, bundleId, bundleId.Length, out var target);
        if (status != 0)
        {
            throw new System.InvalidOperationException($"AECreateDesc failed with OSStatus {status}.");
        }

        int disposeStatus;
        try
        {
            // Wildcards check all events. Zero is the native Boolean false: never request consent.
            status = AEDeterminePermissionToAutomateTarget(ref target, 0x2A2A2A2A, 0x2A2A2A2A, 0);
        }
        finally
        {
            disposeStatus = AEDisposeDesc(ref target);
        }
        if (disposeStatus != 0)
        {
            throw new System.InvalidOperationException($"AEDisposeDesc failed with OSStatus {disposeStatus}.");
        }
        return status;
    }

    [DllImport(AppleEvents)]
    private static extern int AECreateDesc(uint descriptorType, byte[] data, nint dataSize, out Descriptor result);

    [DllImport(AppleEvents)]
    private static extern int AEDeterminePermissionToAutomateTarget(
        ref Descriptor target, uint eventClass, uint eventId, byte askUserIfNeeded);

    [DllImport(AppleEvents)]
    private static extern int AEDisposeDesc(ref Descriptor descriptor);
}
