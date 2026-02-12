namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Vba;

using System.Runtime.Versioning; // Added for platform-specific registry access attribute

/// <summary>
/// Integration tests for VBA Trust Detection functionality.
/// These tests validate VBA trust detection, guidance generation, and TestVbaTrustScope helper.
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
public partial class VbaCommandsTests
{

    /// <summary>
    /// Helper method to check VBA trust status via registry
    /// </summary>
    [SupportedOSPlatform("windows")] // Registry APIs are Windows-only; attribute silences CA1416 analyzer warnings
    protected static bool IsVbaTrustEnabled()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Excel\Security");
            var value = key?.GetValue("AccessVBOM");
            return value != null && (int)value == 1;
        }
        catch (Exception)
        {
            // Test helper - registry access may fail
            return false;
        }
    }
}





