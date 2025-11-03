using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Commands.Table;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture that creates ONE Table test file per test CLASS.
/// The fixture initialization IS the test for Table creation.
/// - Created ONCE before any tests run (~5-10s)
/// - Shared READ-ONLY by all tests in the class
/// - Each test gets its own batch (isolation at batch level)
/// - No file sharing between test classes
/// - Creation results exposed for validation tests
/// </summary>
public class TableTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    
    /// <summary>
    /// Path to the test Table file
    /// </summary>
    public string TestFilePath { get; private set; } = null!;
    
    /// <summary>
    /// Results of Table creation (exposed for validation)
    /// </summary>
    public TableCreationResult CreationResult { get; private set; } = null!;

    public TableTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"TableTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// This IS the test for Table creation - if it fails, all tests fail (correct behavior).
    /// Tests: file creation, data creation, TableCommands.CreateAsync(), persistence.
    /// </summary>
    public async Task InitializeAsync()
    {
        Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        Console.WriteLine("TESTING: Table Creation (via fixture initialization)");
        Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
        
        var sw = Stopwatch.StartNew();
        
        TestFilePath = Path.Join(_tempDir, "TableTest.xlsx");
        CreationResult = new TableCreationResult();
        
        try
        {
            // ═══════════════════════════════════════════════════════
            // TEST 1: File Creation
            // ═══════════════════════════════════════════════════════
            Console.WriteLine("  [1/3] Testing: File creation...");
            var fileCommands = new FileCommands();
            var createFileResult = await fileCommands.CreateEmptyAsync(TestFilePath);
            if (!createFileResult.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: File creation failed: {createFileResult.ErrorMessage}");
            
            CreationResult.FileCreated = true;
            Console.WriteLine("        ✅ File created successfully");
            
            await using var batch = await ExcelSession.BeginBatchAsync(TestFilePath);
            
            // ═══════════════════════════════════════════════════════
            // TEST 2: Data Creation and Table Creation
            // ═══════════════════════════════════════════════════════
            Console.WriteLine("  [2/3] Testing: Data and Table creation...");
            
            // Create sample sales data
            await batch.Execute<int>((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                sheet.Name = "Sales";

                // Add headers
                sheet.Range["A1"].Value2 = "Region";
                sheet.Range["B1"].Value2 = "Product";
                sheet.Range["C1"].Value2 = "Amount";
                sheet.Range["D1"].Value2 = "Date";

                // Add sample data
                sheet.Range["A2"].Value2 = "North";
                sheet.Range["B2"].Value2 = "Widget";
                sheet.Range["C2"].Value2 = 100;
                sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

                sheet.Range["A3"].Value2 = "South";
                sheet.Range["B3"].Value2 = "Gadget";
                sheet.Range["C3"].Value2 = 250;
                sheet.Range["D3"].Value2 = new DateTime(2025, 2, 20);

                sheet.Range["A4"].Value2 = "East";
                sheet.Range["B4"].Value2 = "Widget";
                sheet.Range["C4"].Value2 = 150;
                sheet.Range["D4"].Value2 = new DateTime(2025, 3, 10);

                sheet.Range["A5"].Value2 = "West";
                sheet.Range["B5"].Value2 = "Gadget";
                sheet.Range["C5"].Value2 = 300;
                sheet.Range["D5"].Value2 = new DateTime(2025, 1, 25);

                return 0;
            });
            
            // Create Table using TableCommands
            var tableCommands = new TableCommands();
            var createTableResult = await tableCommands.CreateAsync(
                batch, "Sales", "SalesTable", "A1:D5", hasHeaders: true, tableStyle: "TableStyleMedium2");
                
            if (!createTableResult.Success)
                throw new InvalidOperationException(
                    $"CREATION TEST FAILED: Table creation failed: {createTableResult.ErrorMessage}");
                
            CreationResult.TablesCreated = 1;
            Console.WriteLine("        ✅ Table 'SalesTable' created with 4 rows of data");
            
            // ═══════════════════════════════════════════════════════
            // TEST 3: Persistence (Save)
            // ═══════════════════════════════════════════════════════
            Console.WriteLine("  [3/3] Testing: Batch.SaveAsync() persistence...");
            await batch.SaveAsync();
            Console.WriteLine("        ✅ Table saved successfully");
            
            sw.Stop();
            CreationResult.Success = true;
            CreationResult.CreationTimeSeconds = sw.Elapsed.TotalSeconds;
            
            Console.WriteLine();
            Console.WriteLine("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
            Console.WriteLine($"✅ CREATION TEST PASSED in {sw.Elapsed.TotalSeconds:F1}s");
            Console.WriteLine($"   📊 {CreationResult.TablesCreated} table created with sample data");
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
}

/// <summary>
/// Results of Table creation (exposed by fixture for validation tests)
/// </summary>
public class TableCreationResult
{
    public bool Success { get; set; }
    public bool FileCreated { get; set; }
    public int TablesCreated { get; set; }
    public double CreationTimeSeconds { get; set; }
    public string? ErrorMessage { get; set; }
}
