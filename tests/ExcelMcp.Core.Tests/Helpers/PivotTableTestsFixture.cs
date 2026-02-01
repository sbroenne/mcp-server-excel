using System.Diagnostics;
using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

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
    public Task InitializeAsync()
    {
        var sw = Stopwatch.StartNew();

        TestFilePath = Path.Join(_tempDir, "PivotTableTest.xlsx");
        CreationResult = new PivotTableCreationResult();

        try
        {
            using (var manager = new SessionManager())
            {
                var sessionId = manager.CreateSessionForNewFile(TestFilePath, showExcel: false);
                manager.CloseSession(sessionId, save: true);
            }

            CreationResult.FileCreated = true;

            using var batch = ExcelSession.BeginBatch(TestFilePath);
            batch.Execute((ctx, ct) =>
            {
                dynamic sheet = ctx.Book.Worksheets.Item(1);
                sheet.Name = "SalesData";

                // Headers
                sheet.Range["A1"].Value2 = "Region";
                sheet.Range["B1"].Value2 = "Product";
                sheet.Range["C1"].Value2 = "Sales";
                sheet.Range["D1"].Value2 = "Date";

                // Sample data (5 rows)
                sheet.Range["A2"].Value2 = "North";
                sheet.Range["B2"].Value2 = "Widget";
                sheet.Range["C2"].Value2 = 100;
                sheet.Range["D2"].Value2 = new DateTime(2025, 1, 15);

                sheet.Range["A3"].Value2 = "North";
                sheet.Range["B3"].Value2 = "Widget";
                sheet.Range["C3"].Value2 = 150;
                sheet.Range["D3"].Value2 = new DateTime(2025, 1, 20);

                sheet.Range["A4"].Value2 = "South";
                sheet.Range["B4"].Value2 = "Gadget";
                sheet.Range["C4"].Value2 = 200;
                sheet.Range["D4"].Value2 = new DateTime(2025, 2, 10);

                sheet.Range["A5"].Value2 = "North";
                sheet.Range["B5"].Value2 = "Gadget";
                sheet.Range["C5"].Value2 = 75;
                sheet.Range["D5"].Value2 = new DateTime(2025, 2, 15);

                sheet.Range["A6"].Value2 = "South";
                sheet.Range["B6"].Value2 = "Widget";
                sheet.Range["C6"].Value2 = 125;
                sheet.Range["D6"].Value2 = new DateTime(2025, 3, 5);

                return 0;
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

            throw; // Fail all tests in class (correct behavior - no point testing if creation failed)
        }

        return Task.CompletedTask;
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
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(filePath, showExcel: false);
        manager.CloseSession(sessionId, save: true);
        return filePath;
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
