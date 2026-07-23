// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture for Chart tests — creates ONE shared workbook with pre-populated data.
/// All tests share this file safely because ExcelBatch.Dispose() closes WITHOUT saving.
///
/// Performance: Creates 2 COM sessions at startup instead of 2 per test (74 tests = 148→2 sessions saved).
/// Data layout on Sheet1: A1:D6 with 4 columns × 5 data rows (covers most test scenarios).
/// </summary>
public class ChartTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Temp directory for all test files (auto-cleaned on disposal)
    /// </summary>
    public string TempDir => _tempDir;

    /// <summary>
    /// Pre-populated test file shared by all tests. Safe to share because
    /// ExcelBatch.Dispose() closes the workbook WITHOUT saving — each test
    /// sees the original saved state (data, no charts).
    /// </summary>
    public string SharedTestFile { get; private set; } = null!;

    public ChartTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"ChartTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Creates a copy of the shared test file for tests that need file isolation.
    /// Uses File.Copy (microseconds) instead of COM session (seconds).
    /// </summary>
    public string CreateTestFile([CallerMemberName] string testName = "", string extension = ".xlsx")
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}{extension}";
        var filePath = Path.Join(_tempDir, fileName);
        File.Copy(SharedTestFile, filePath);
        return filePath;
    }

    /// <summary>
    /// Called ONCE before any tests run. Creates shared workbook with standard data.
    /// </summary>
    public Task InitializeAsync()
    {
        SharedTestFile = Path.Join(_tempDir, "SharedChartTestFile.xlsx");

        // Create file — 1st COM session
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(SharedTestFile, show: false);
        manager.CloseSession(sessionId, save: true);

        // Pre-populate standard data — 2nd COM session
        using var batch = ExcelSession.BeginBatch(SharedTestFile);
        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets[1];
            // 4 columns × 5 data rows — covers A1:B3, A1:B4, A1:B5, A1:C4, A1:D4 patterns
            sheet.Range["A1:D6"].Value2 = new object[,]
            {
                { "X", "Y", "Series2", "Series3" },
                { 1, 100, 15, 50 },
                { 2, 150, 25, 75 },
                { 3, 200, 35, 100 },
                { 4, 250, 45, 125 },
                { 5, 300, 55, 150 }
            };
            return 0;
        });
        batch.Save();

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
}




