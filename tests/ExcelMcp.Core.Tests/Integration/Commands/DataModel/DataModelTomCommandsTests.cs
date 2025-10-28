using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.DataModel;

/// <summary>
/// Integration tests for Data Model TOM (Tabular Object Model) operations.
/// These tests require Excel installation and validate TOM Data Model operations.
/// Tests use Core commands directly (not through CLI wrapper).
/// </summary>
[Trait("Layer", "Core")]
[Trait("Category", "Integration")]
[Trait("RequiresExcel", "true")]
[Trait("Feature", "DataModelTom")]
public partial class DataModelTomCommandsTests : IDisposable
{
    private readonly IDataModelTomCommands _tomCommands;
    private readonly IDataModelCommands _dataModelCommands;
    private readonly IFileCommands _fileCommands;
    private readonly string _testExcelFile;
    private readonly string _tempDir;
    private bool _disposed;

    public DataModelTomCommandsTests()
    {
        _tomCommands = new DataModelTomCommands();
        _dataModelCommands = new DataModelCommands();
        _fileCommands = new FileCommands();

        // Create temp directory for test files
        _tempDir = Path.Combine(Path.GetTempPath(), $"ExcelCore_DM_TOM_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);

        _testExcelFile = Path.Combine(_tempDir, "TestDataModelTom.xlsx");

        // Create test Excel file with Data Model
        CreateTestDataModelFile();
    }

    private void CreateTestDataModelFile()
    {
        // Use Task.Run() to properly execute async method synchronously
        var task = Task.Run(async () =>
        {
            // Create an empty workbook first
            var result = await _fileCommands.CreateEmptyAsync(_testExcelFile, overwriteIfExists: false);
            if (!result.Success)
            {
                throw new InvalidOperationException($"Failed to create test Excel file: {result.ErrorMessage}. Excel may not be installed.");
            }

            // Create realistic Data Model with sample data
            try
            {
                await DataModelTestHelper.CreateSampleDataModelAsync(_testExcelFile);
            }
            catch (Exception ex)
            {
                // Data Model creation may fail on some Excel versions
                // Tests will handle this gracefully by checking for "no Data Model" errors
                System.Diagnostics.Debug.WriteLine($"Could not create sample Data Model: {ex.Message}");
            }
        });

        task.GetAwaiter().GetResult();
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
