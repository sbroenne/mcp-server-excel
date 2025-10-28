using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Base class for Data Model Core operations integration tests.
/// These tests require Excel installation and validate Core Data Model operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModel")]
public partial class DataModelCommandsTests : IDisposable
{
    protected readonly IDataModelCommands _dataModelCommands;
    protected readonly IFileCommands _fileCommands;
    protected readonly string _testExcelFile;
    protected readonly string _testMeasureFile;
    protected readonly string _tempDir;
    private bool _disposed;

    public DataModelCommandsTests()
    {
        _dataModelCommands = new DataModelCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_DM_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestDataModel.xlsx");
        _testMeasureFile = Path.Combine(_tempDir, "TestMeasure.dax");

        // Create test Excel file with Data Model
        CreateTestDataModelFile();
    }

    private void CreateTestDataModelFile()
    {
        // Create an empty workbook first (synchronous helper)
        var task = Task.Run(async () =>
            await _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false));
        var result = task.GetAwaiter().GetResult();

        if (!result.Success)
        {
            throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
        }

        // Create realistic Data Model with sample data
        try
        {
            var createTask = Task.Run(async () =>
                await DataModelTestHelper.CreateSampleDataModelAsync(_testExcelFile));
            createTask.GetAwaiter().GetResult();
        }
        catch (Exception ex)
        {
            // Data Model creation may fail on some Excel versions
            // Tests will handle this gracefully by checking for "no Data Model" errors
            System.Diagnostics.Debug.WriteLine($"Could not create sample Data Model: {ex.Message}");
        }
    }

    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(_tempDir))
            {
                // Give Excel time to release file locks
                System.Threading.Thread.Sleep(100);

                // Retry cleanup a few times if needed
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Directory.Delete(_tempDir, recursive: true);
                        break;
                    }
                    catch (IOException) when (i < 2)
                    {
                        System.Threading.Thread.Sleep(500);
                    }
                }
            }
        }
        catch
        {
            // Best effort cleanup
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
