// <copyright file="SheetTestsFixture.cs" company="Stephan Brenner">
// Copyright (c) Stephan Brenner. All rights reserved.
// </copyright>

using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Shared fixture for SheetCommandsTests that provides a single Excel file for single-workbook tests.
/// Cross-file tests (CopyToFile, MoveToFile) create their own file pairs.
/// Uses IAsyncLifetime to create file once for all tests, reducing test execution time.
/// </summary>
public class SheetTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    private int _sheetCounter;

    /// <summary>
    /// Initializes a new instance of the <see cref="SheetTestsFixture"/> class.
    /// </summary>
    public SheetTestsFixture()
    {
        _tempDir = Path.Combine(Path.GetTempPath(), $"SheetTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Gets the path to the shared test file for single-workbook tests.
    /// </summary>
    public string TestFilePath { get; private set; } = string.Empty;

    /// <summary>
    /// Gets the temp directory for cross-workbook tests that need their own files.
    /// </summary>
    public string TempDir => _tempDir;

    /// <summary>
    /// Creates the shared Excel file.
    /// </summary>
    public Task InitializeAsync()
    {
        TestFilePath = Path.Combine(_tempDir, "SheetTests_Shared.xlsx");
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(TestFilePath, show: false);
        manager.CloseSession(sessionId, save: true);
        return Task.CompletedTask;
    }

    /// <summary>
    /// Cleans up the temp directory and all test files.
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

    /// <summary>
    /// Creates a unique sheet name for test isolation within the shared workbook.
    /// Call this at the start of each single-workbook test to get a unique sheet to work with.
    /// </summary>
    /// <param name="batch">The Excel batch to create the sheet in.</param>
    /// <param name="testName">Test method name (auto-captured via CallerMemberName).</param>
    /// <returns>The name of the created test sheet.</returns>
    public string CreateTestSheet(IExcelBatch batch, [CallerMemberName] string testName = "")
    {
        var sheetNum = Interlocked.Increment(ref _sheetCounter);
        var sheetName = $"T{sheetNum}_{testName}";

        // Truncate if too long (Excel max is 31 chars)
        if (sheetName.Length > 31)
        {
            sheetName = $"T{sheetNum}_{testName[..(31 - $"T{sheetNum}_".Length)]}";
        }

        batch.Execute((ctx, ct) =>
        {
            dynamic sheet = ctx.Book.Worksheets.Add();
            sheet.Name = sheetName;
        });

        return sheetName;
    }

    /// <summary>
    /// Creates a unique test file for cross-workbook tests that need their own isolated files.
    /// </summary>
    /// <param name="testName">Test method name.</param>
    /// <param name="suffix">Optional suffix to distinguish source/target files.</param>
    /// <returns>Path to the created test file.</returns>
    public string CreateCrossWorkbookTestFile(string testName, string suffix = "")
    {
        var fileName = string.IsNullOrEmpty(suffix)
            ? $"{testName}_{Guid.NewGuid():N}.xlsx"
            : $"{testName}_{suffix}_{Guid.NewGuid():N}.xlsx";

        // Truncate filename if too long
        if (fileName.Length > 200)
        {
            fileName = $"{testName[..50]}_{suffix}_{Guid.NewGuid():N}.xlsx";
        }

        var filePath = Path.Combine(_tempDir, fileName);
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(filePath, show: false);
        manager.CloseSession(sessionId, save: true);
        return filePath;
    }
}




