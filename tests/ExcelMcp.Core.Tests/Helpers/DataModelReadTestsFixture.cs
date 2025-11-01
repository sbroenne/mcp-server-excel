using System;
using System.IO;
using System.Threading.Tasks;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Shared fixture for Data Model READ tests.
/// Copies pre-built DataModelTemplate.xlsx for fast setup (~0.5s vs 60-120s build time).
/// Each test class gets its own copy of the template.
/// </summary>
public class DataModelReadTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    
    /// <summary>
    /// Path to the test Data Model file (copy of template)
    /// </summary>
    public string TestFilePath { get; private set; } = null!;

    public DataModelReadTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"DataModelReadTests_{Guid.NewGuid():N}");
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// Copies the pre-built template (FAST - ~0.5s vs 60-120s creation).
    /// </summary>
    public Task InitializeAsync()
    {
        Directory.CreateDirectory(_tempDir);
        TestFilePath = Path.Join(_tempDir, "DataModel.xlsx");

        // Find template relative to test assembly
        var testAssemblyDir = AppContext.BaseDirectory;
        var solutionRoot = Path.GetFullPath(Path.Join(testAssemblyDir, "../../../../.."));
        var templatePath = Path.Join(solutionRoot, "tests/ExcelMcp.Core.Tests/TestAssets/DataModelTemplate.xlsx");

        if (!File.Exists(templatePath))
        {
            throw new FileNotFoundException(
                $"Data Model template not found. Generate it by running:\n" +
                $"  dotnet test tests/ExcelMcp.Core.Tests --filter \"FullyQualifiedName~BuildDataModelAsset\"\n" +
                $"Expected path: {templatePath}");
        }

        // Copy template (fast!)
        File.Copy(templatePath, TestFilePath, overwrite: true);
        
        return Task.CompletedTask;
    }

    /// <summary>
    /// Called ONCE after all tests in the class complete.
    /// </summary>
    public Task DisposeAsync()
    {
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
        return Task.CompletedTask;
    }
}
