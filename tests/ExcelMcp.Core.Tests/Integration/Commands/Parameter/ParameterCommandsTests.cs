using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.Parameter;

/// <summary>
/// Integration tests for Parameter Core operations using Excel COM automation.
/// Tests Core layer directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Fast")]
[Trait("Feature", "Parameters")]
[Trait("RequiresExcel", "true")]
public partial class ParameterCommandsTests : IDisposable
{
    private readonly IParameterCommands _parameterCommands;
    private readonly string _tempDir;
    private bool _disposed;

    public ParameterCommandsTests()
    {
        _parameterCommands = new ParameterCommands();
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_ParamTests_{Guid.NewGuid():N}");
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
