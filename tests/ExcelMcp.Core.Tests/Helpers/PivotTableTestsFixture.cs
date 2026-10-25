using System.Diagnostics;
using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture that creates ONE PivotTable test file per test CLASS.
/// The fixture initialization IS the test for PivotTable data preparation.
/// - Created ONCE before any tests run (~5-10s)
/// - Shared READ-ONLY by all tests in the class
/// - Each test gets its own batch (isolation at batch level)
/// - No file sharing between test classes
/// - Creation results exposed for validation tests
/// - CreateTestFile() available for tests that need unique files
/// </summary>
public class PivotTableTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Temp directory for all test files (auto-cleaned on disposal)
    /// </summary>
    public string TempDir => _tempDir;

    /// <summary>
    /// Path to the test PivotTable file
    /// </summary>
    public string TestFilePath { get; private set; } = null!;

    /// <summary>
    /// Results of data creation (exposed for validation)
    /// </summary>
    public PivotTableCreationResult CreationResult { get; private set; } = null!;

    public PivotTableTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"PivotTableTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// This IS the test for data preparation - if it fails, all tests fail (correct behavior).
    /// Tests: file creation, sales data creation, persistence.
    /// </summary>
    public async Task InitializeAsync()
    {
        await Task.Run(() => InitializeCore());
    }

    private void InitializeCore()
    {
        var sw = Stopwatch.StartNew();

        TestFilePath = Path.Join(_tempDir, "PivotTableTest.xlsx");
        CreationResult = new PivotTableCreationResult();

        try
        {
            SavedWorkbookTemplates.CopyBlankTo(TestFilePath);

            CreationResult.FileCreated = true;

            using var batch = ExcelSession.BeginBatch(TestFilePath);
            batch.Execute((ctx, ct) =>
            {
                Excel.Sheets? sheets = null;
                Excel.Worksheet? sheet = null;
                Excel.Range? range = null;
                try
                {
                    sheets = ctx.Book.Worksheets;
                    sheet = (Excel.Worksheet)sheets[1];
                    sheet.Name = "SalesData";
                    range = sheet.Range["A1:D6"];
                    range.Value2 = new object[,]
                    {
                        { "Region", "Product", "Sales", "Date" },
                        { "North", "Widget", 100, new DateTime(2025, 1, 15) },
                        { "North", "Widget", 150, new DateTime(2025, 1, 20) },
                        { "South", "Gadget", 200, new DateTime(2025, 2, 10) },
                        { "North", "Gadget", 75, new DateTime(2025, 2, 15) },
                        { "South", "Widget", 125, new DateTime(2025, 3, 5) }
                    };
                    return 0;
                }
                finally
                {
                    ComUtilities.Release(ref range);
                    ComUtilities.Release(ref sheet);
                    ComUtilities.Release(ref sheets);
                }
            });

            CreationResult.DataRowsCreated = 5;

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

            throw; // Fail all tests in class
        }
    }

    /// <summary>
    /// Creates a unique test file for tests that need their own file.
    /// File name includes test name + GUID for uniqueness.
    /// </summary>
    /// <param name="testName">Test name (auto-populated via CallerMemberName)</param>
    /// <param name="extension">File extension (default: .xlsx)</param>
    /// <returns>Path to the new test file</returns>
    public string CreateTestFile([CallerMemberName] string testName = "", string extension = ".xlsx")
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}{extension}";
        var filePath = Path.Join(_tempDir, fileName);
        return SavedWorkbookTemplates.CopyBlankTo(filePath);
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
/// Results of PivotTable data creation (exposed by fixture for validation tests)
/// </summary>
public class PivotTableCreationResult
{
    /// <inheritdoc/>
    public bool Success { get; set; }
    /// <inheritdoc/>
    public bool FileCreated { get; set; }
    /// <inheritdoc/>
    public int DataRowsCreated { get; set; }
    /// <inheritdoc/>
    public double CreationTimeSeconds { get; set; }
    /// <inheritdoc/>
    public string? ErrorMessage { get; set; }
}

