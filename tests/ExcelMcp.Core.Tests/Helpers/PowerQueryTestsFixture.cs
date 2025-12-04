using System.Diagnostics;
using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture that creates ONE Power Query test file per test CLASS.
/// The fixture initialization IS the test for Power Query creation.
/// - Created ONCE before any tests run (~10-15s)
/// - Shared READ-ONLY by all tests in the class
/// - Each test gets its own batch (isolation at batch level)
/// - No file sharing between test classes
/// - Creation results exposed for validation tests
/// - CreateTestFile() available for tests that need unique files
/// </summary>
public class PowerQueryTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    private readonly FileCommands _fileCommands = new();

    /// <summary>
    /// Temp directory for all test files (auto-cleaned on disposal)
    /// </summary>
    public string TempDir => _tempDir;

    /// <summary>
    /// Path to the test Power Query file
    /// </summary>
    public string TestFilePath { get; private set; } = null!;

    /// <summary>
    /// Results of Power Query creation (exposed for validation)
    /// </summary>
    public PowerQueryCreationResult CreationResult { get; private set; } = null!;

    public PowerQueryTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"PowerQueryTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// This IS the test for Power Query creation - if it fails, all tests fail (correct behavior).
    /// Tests: file creation, M code file creation, Import, persistence.
    /// </summary>
    public Task InitializeAsync()
    {
        var sw = Stopwatch.StartNew();

        TestFilePath = Path.Join(_tempDir, "PowerQuery.xlsx");
        CreationResult = new PowerQueryCreationResult();

        try
        {
            var fileCommands = new FileCommands();
            fileCommands.CreateEmpty(TestFilePath);

            CreationResult.FileCreated = true;

            using var batch = ExcelSession.BeginBatch(TestFilePath);

            var mCodeFiles = new string[3];
            mCodeFiles[0] = CreateMCodeFile("BasicQuery", CreateBasicMCode());
            mCodeFiles[1] = CreateMCodeFile("DataQuery", CreateDataQueryMCode());
            mCodeFiles[2] = CreateMCodeFile("RefreshableQuery", CreateRefreshableQueryMCode());
            CreationResult.MCodeFilesCreated = 3;

            var dataModelCommands = new DataModelCommands();
            var powerQueryCommands = new PowerQueryCommands(dataModelCommands);

            // Create throws on error, so reaching here means success
            powerQueryCommands.Create(batch, "BasicQuery", File.ReadAllText(mCodeFiles[0]), PowerQueryLoadMode.ConnectionOnly);
            powerQueryCommands.Create(batch, "DataQuery", File.ReadAllText(mCodeFiles[1]), PowerQueryLoadMode.ConnectionOnly);
            powerQueryCommands.Create(batch, "RefreshableQuery", File.ReadAllText(mCodeFiles[2]), PowerQueryLoadMode.ConnectionOnly);

            CreationResult.QueriesImported = 3;

            // ═══════════════════════════════════════════════════════
            // TEST 4: Persistence (Save)
            // ═══════════════════════════════════════════════════════
            batch.Save();

            sw.Stop();
            CreationResult.Success = true;
            CreationResult.CreationTimeSeconds = sw.Elapsed.TotalSeconds;

        }
        catch (Exception ex)
        {
            CreationResult.Success = false;
            CreationResult.ErrorMessage = ex.Message;

            sw.Stop();

            throw; // Fail all tests in class (correct behavior - no point testing if creation failed)
        }

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
            // Cleanup is best-effort
        }
        return Task.CompletedTask;
    }

    /// <summary>
    /// Creates a unique empty test file for the calling test.
    /// Uses [CallerMemberName] to auto-populate the test name.
    /// </summary>
    /// <param name="testName">Auto-populated from caller method name</param>
    /// <returns>Path to the unique test file</returns>
    public string CreateTestFile([CallerMemberName] string testName = "")
    {
        var guid = Guid.NewGuid().ToString("N")[..8];
        var testFile = Path.Join(_tempDir, $"PQ_{testName}_{guid}.xlsx");
        _fileCommands.CreateEmpty(testFile);
        return testFile;
    }

    /// <summary>
    /// Creates an M code file in the temp directory
    /// </summary>
    private string CreateMCodeFile(string name, string mCode)
    {
        var filePath = Path.Join(_tempDir, $"{name}.pq");
        File.WriteAllText(filePath, mCode);
        return filePath;
    }

    /// <summary>
    /// Creates basic M code for simple queries
    /// </summary>
    private static string CreateBasicMCode()
    {
        return @"let
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
    }

    /// <summary>
    /// Creates M code with more data for testing
    /// </summary>
    private static string CreateDataQueryMCode()
    {
        return @"let
    Source = #table(
        {""ID"", ""Name"", ""Value""},
        {
            {1, ""Item1"", 100},
            {2, ""Item2"", 200},
            {3, ""Item3"", 300},
            {4, ""Item4"", 400},
            {5, ""Item5"", 500}
        }
    )
in
    Source";
    }

    /// <summary>
    /// Creates M code for refreshable query testing
    /// </summary>
    private static string CreateRefreshableQueryMCode()
    {
        return @"let
    Source = #table(
        {""Date"", ""Amount""},
        {
            {#date(2024, 1, 1), 1000},
            {#date(2024, 2, 1), 2000},
            {#date(2024, 3, 1), 3000}
        }
    )
in
    Source";
    }
}

/// <summary>
/// Results of Power Query creation (exposed by fixture for validation tests)
/// </summary>
public class PowerQueryCreationResult
{
    /// <inheritdoc/>
    public bool Success { get; set; }
    /// <inheritdoc/>
    public bool FileCreated { get; set; }
    /// <inheritdoc/>
    public int MCodeFilesCreated { get; set; }
    /// <inheritdoc/>
    public int QueriesImported { get; set; }
    /// <inheritdoc/>
    public double CreationTimeSeconds { get; set; }
    /// <inheritdoc/>
    public string? ErrorMessage { get; set; }
}
