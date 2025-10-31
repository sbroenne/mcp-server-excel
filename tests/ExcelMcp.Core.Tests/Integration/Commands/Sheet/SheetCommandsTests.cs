using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Sheet;

/// <summary>
/// Integration tests for Sheet lifecycle operations.
/// These tests require Excel installation and validate Core worksheet lifecycle management.
/// Tests use Core commands directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
/// Data operations (read, write, clear) moved to RangeCommandsTests.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "Worksheets")]
public partial class SheetCommandsTests : IDisposable
{
    private readonly ISheetCommands _sheetCommands;
    private readonly string _tempDir;
    private bool _disposed;

    public SheetCommandsTests()
    {
        _sheetCommands = new SheetCommands();
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_Sheet_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
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
            // Ignore cleanup failures
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
