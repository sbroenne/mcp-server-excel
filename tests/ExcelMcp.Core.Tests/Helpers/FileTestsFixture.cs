using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture for File tests providing efficient test file creation.
/// Each test gets a unique .xlsx file via CreateTestFile() method.
/// </summary>
public class FileTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Temp directory for all test files (auto-cleaned on disposal)
    /// </summary>
    public string TempDir => _tempDir;

    public FileTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"FileTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <inheritdoc/>
    public Task InitializeAsync()
    {
        return Task.CompletedTask;
    }

    /// <inheritdoc/>
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
    /// Creates a unique empty .xlsx test file for the calling test.
    /// Uses [CallerMemberName] to auto-populate the test name.
    /// </summary>
    /// <param name="testName">Auto-populated from caller method name</param>
    /// <returns>Path to the unique .xlsx test file</returns>
    public string CreateTestFile([CallerMemberName] string testName = "")
    {
        var guid = Guid.NewGuid().ToString("N")[..8];
        var testFile = Path.Join(_tempDir, $"File_{testName}_{guid}.xlsx");
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(testFile, showExcel: false);
        manager.CloseSession(sessionId, save: true);
        return testFile;
    }
}
