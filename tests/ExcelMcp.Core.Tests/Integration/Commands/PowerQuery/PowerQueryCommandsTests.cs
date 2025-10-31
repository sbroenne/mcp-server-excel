using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.PowerQuery;

/// <summary>
/// Integration tests for Power Query Core operations.
/// These tests require Excel installation and validate Core Power Query data operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// Each test uses a unique Excel file for complete test isolation.
///
/// For comprehensive workflow tests (mode switching), see PowerQueryLoadConfigWorkflowTests.cs.
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "PowerQuery")]
public partial class PowerQueryCommandsTests : IDisposable
{
    protected readonly IPowerQueryCommands _powerQueryCommands;
    protected readonly IFileCommands _fileCommands;
    protected readonly ISheetCommands _sheetCommands;
    protected readonly string _tempDir;
    private bool _disposed;

    public PowerQueryCommandsTests()
    {
        var dataModelCommands = new DataModelCommands();
        _powerQueryCommands = new PowerQueryCommands(dataModelCommands);
        _fileCommands = new FileCommands();
        _sheetCommands = new SheetCommands();

        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_PQ_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Creates a unique Excel file for a test to avoid parallel execution conflicts.
    /// Each test gets its own isolated file.
    /// </summary>
    protected string CreateUniqueTestExcelFile(string testName)
    {
        var uniqueFile = Path.Combine(_tempDir, $"{testName}_{Guid.NewGuid():N}.xlsx");
        var result = _fileCommands.CreateEmptyAsync(uniqueFile, overwriteIfExists: false).GetAwaiter().GetResult();
        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }
        return uniqueFile;
    }

    /// <summary>
    /// Creates a unique test Power Query M code file.
    /// Each test gets its own isolated M code file.
    /// </summary>
    protected string CreateUniqueTestQueryFile(string testName)
    {
        var uniqueFile = Path.Combine(_tempDir, $"{testName}_{Guid.NewGuid():N}.pq");
        string mCode = @"let
    Source = #table(
        {""Column1"", ""Column2"", ""Column3""},
        {
            {""Value1"", ""Value2"", ""Value3""},
            {""A"", ""B"", ""C""},
            {""X"", ""Y"", ""Z""}
        }
    )
in
    Source";

        File.WriteAllText(uniqueFile, mCode);
        return uniqueFile;
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
