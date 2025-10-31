using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.FileOperations;

/// <summary>
/// Integration tests for File Core operations using Excel COM automation.
/// Tests Core layer directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
/// 
/// WHAT LLMs NEED TO KNOW:
/// 1. CreateEmpty creates .xlsx or .xlsm files (valid extensions)
/// 2. CreateEmpty fails on invalid extensions (.xls, .csv, .txt, etc.)
/// 3. CreateEmpty respects overwrite flag (default: fail if exists)
/// 4. TestFile validates existence and extension
/// 5. Result objects have Success, ErrorMessage, FilePath properties
///
/// LAYER RESPONSIBILITY:
/// - ✅ Test Excel COM file operations and Result objects
/// - ✅ Test business rules (valid extensions, overwrite behavior)
/// - ❌ DO NOT test CLI argument parsing (CLI's responsibility)
/// - ❌ DO NOT test JSON serialization (MCP Server's responsibility)
/// - ❌ DO NOT test infrastructure (paths, directories, OS validation)
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("Speed", "Medium")]
[Trait("Feature", "Files")]
[Trait("RequiresExcel", "true")]
public partial class FileCommandsTests : IDisposable
{
    private readonly IFileCommands _fileCommands;
    private readonly string _tempDir;
    private bool _disposed;

    public FileCommandsTests()
    {
        _fileCommands = new FileCommands();
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_File_Tests_{Guid.NewGuid():N}");
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
