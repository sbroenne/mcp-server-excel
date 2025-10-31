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
public partial class VbaTrustDetectionTests : IDisposable
{
    private readonly IScriptCommands _scriptCommands;
    private readonly string _tempDir;
    private bool _disposed;

    public VbaTrustDetectionTests()
    {
        _scriptCommands = new ScriptCommands();
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_VBATrust_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
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
        catch
        {
            return false;
        }
    }

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                Directory.Delete(_tempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup failures shouldn't fail tests
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
