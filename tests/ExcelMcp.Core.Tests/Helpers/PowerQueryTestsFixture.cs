using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
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
/// </summary>
public class PowerQueryTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    
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
    /// Tests: file creation, M code file creation, ImportAsync, persistence.
    /// </summary>
    public async Task InitializeAsync()
    {
        Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        Console.WriteLine("TESTING: Power Query Creation (via fixture initialization)");
        Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        
        var sw = Stopwatch.StartNew();
        
        TestFilePath = Path.Join(_tempDir, "PowerQuery.xlsx");
        CreationResult = new PowerQueryCreationResult();
        
        try
        {
            // ═══════════════════════════════════════════════════════
            // TEST 1: File Creation
            // ═══════════════════════════════════════════════════════
            Console.WriteLine("  [1/4] Testing: File creation...");
            var fileCommands = new FileCommands();
            var createFileResult = await fileCommands.CreateEmptyAsync(TestFilePath);
            if (!createFileResult.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: File creation failed: {createFileResult.ErrorMessage}");
            
            CreationResult.FileCreated = true;
            Console.WriteLine("        ✅ File created successfully");
            
            await using var batch = await ExcelSession.BeginBatchAsync(TestFilePath);
            
            // ═══════════════════════════════════════════════════════
            // TEST 2: M Code File Creation
            // ═══════════════════════════════════════════════════════
            Console.WriteLine("  [2/4] Testing: M code file creation...");
            var mCodeFiles = new string[3];
            mCodeFiles[0] = CreateMCodeFile("BasicQuery", CreateBasicMCode());
            mCodeFiles[1] = CreateMCodeFile("DataQuery", CreateDataQueryMCode());
            mCodeFiles[2] = CreateMCodeFile("RefreshableQuery", CreateRefreshableQueryMCode());
            CreationResult.MCodeFilesCreated = 3;
            Console.WriteLine("        ✅ Created 3 M code files");
            
            // ═══════════════════════════════════════════════════════
            // TEST 3: PowerQueryCommands.ImportAsync()
            // ═══════════════════════════════════════════════════════
            Console.WriteLine("  [3/4] Testing: PowerQueryCommands.ImportAsync() for 3 queries...");
            var dataModelCommands = new DataModelCommands();
            var powerQueryCommands = new PowerQueryCommands(dataModelCommands);
            
            var import1 = await powerQueryCommands.ImportAsync(batch, "BasicQuery", mCodeFiles[0]);
            if (!import1.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: ImportAsync(BasicQuery) failed: {import1.ErrorMessage}");
                
            var import2 = await powerQueryCommands.ImportAsync(batch, "DataQuery", mCodeFiles[1]);
            if (!import2.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: ImportAsync(DataQuery) failed: {import2.ErrorMessage}");
                
            var import3 = await powerQueryCommands.ImportAsync(batch, "RefreshableQuery", mCodeFiles[2]);
            if (!import3.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: ImportAsync(RefreshableQuery) failed: {import3.ErrorMessage}");
                
            CreationResult.QueriesImported = 3;
            Console.WriteLine("        ✅ Imported 3 Power Queries");
            
            // ═══════════════════════════════════════════════════════
            // TEST 4: Persistence (Save)
            // ═══════════════════════════════════════════════════════
            Console.WriteLine("  [4/4] Testing: Batch.SaveAsync() persistence...");
            await batch.SaveAsync();
            Console.WriteLine("        ✅ Power Queries saved successfully");
            
            sw.Stop();
            CreationResult.Success = true;
            CreationResult.CreationTimeSeconds = sw.Elapsed.TotalSeconds;
            
            Console.WriteLine();
            Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            Console.WriteLine($"✅ CREATION TEST PASSED in {sw.Elapsed.TotalSeconds:F1}s");
            Console.WriteLine($"   📄 {CreationResult.MCodeFilesCreated} M code files created");
            Console.WriteLine($"   📊 {CreationResult.QueriesImported} Power Queries imported");
            Console.WriteLine($"   💾 File: {TestFilePath}");
            Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            Console.WriteLine();
        }
        catch (Exception ex)
        {
            CreationResult.Success = false;
            CreationResult.ErrorMessage = ex.Message;
            
            sw.Stop();
            Console.WriteLine();
            Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            Console.WriteLine($"❌ CREATION TEST FAILED after {sw.Elapsed.TotalSeconds:F1}s");
            Console.WriteLine($"   Error: {ex.Message}");
            Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            Console.WriteLine();
            
            throw; // Fail all tests in class (correct behavior - no point testing if creation failed)
        }
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
    /// Creates an M code file in the temp directory
    /// </summary>
    private string CreateMCodeFile(string name, string mCode)
    {
        var filePath = Path.Join(_tempDir, $"{name}.pq");
        System.IO.File.WriteAllText(filePath, mCode);
        return filePath;
    }

    /// <summary>
    /// Creates basic M code for simple queries
    /// </summary>
    private string CreateBasicMCode()
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
    private string CreateDataQueryMCode()
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
    private string CreateRefreshableQueryMCode()
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
    public bool Success { get; set; }
    public bool FileCreated { get; set; }
    public int MCodeFilesCreated { get; set; }
    public int QueriesImported { get; set; }
    public double CreationTimeSeconds { get; set; }
    public string? ErrorMessage { get; set; }
}
