using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.VbaTrust;

/// <summary>
/// Integration tests for VBA Trust Detection functionality.
/// These tests validate VBA trust detection, guidance generation, and TestVbaTrustScope helper.
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "VBATrust")]
public partial class VbaTrustDetectionTests : IClassFixture<TempDirectoryFixture>
{
    private readonly IVbaCommands _scriptCommands;
    private readonly string _tempDir;

    public VbaTrustDetectionTests(TempDirectoryFixture fixture)
    {
        _scriptCommands = new VbaCommands();
        _tempDir = fixture.TempDir;
    }

    /// <summary>
    /// Helper method to check VBA trust status via registry
    /// </summary>
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

