using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Setup;

/// <summary>
/// Integration tests for Setup Core operations.
/// Tests Core layer directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Setup")]
[Trait("RequiresExcel", "true")]
public partial class SetupCommandsTests : IDisposable
{
    private readonly ISetupCommands _setupCommands;
    private readonly string _tempDir;
    private bool _disposed;

    public SetupCommandsTests()
    {
        _setupCommands = new SetupCommands();
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_SetupTests_{Guid.NewGuid():N}");
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
            // Ignore cleanup errors
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
