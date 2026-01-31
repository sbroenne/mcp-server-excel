// Copyright (c) Sbroenne. All rights reserved.
// Licensed under the MIT License.

using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture for Chart tests that need isolated test files.
/// Provides temp directory and file creation helpers.
/// - Created ONCE before any tests run
/// - Each test creates unique files via CreateTestFile()
/// - Temp directory cleaned up after all tests complete
/// </summary>
public class ChartTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Temp directory for all test files (auto-cleaned on disposal)
    /// </summary>
    public string TempDir => _tempDir;

    public ChartTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"ChartTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
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

        // Use SessionManager to create file and immediately close the session
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(filePath, showExcel: false);
        manager.CloseSession(sessionId, save: true);

        return filePath;
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// </summary>
    public Task InitializeAsync()
    {
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
