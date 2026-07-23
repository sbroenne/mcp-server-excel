using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture for Connection tests that provides efficient test file creation.
/// Connection tests often need isolated files because they:
/// - Create/delete connections within a single test
/// - Need external source files for OLEDB connections
/// - Modify connection state
///
/// This fixture provides:
/// - A shared temp directory (auto-cleaned on disposal)
/// - Fast test file creation method
/// - Helper for creating external source files
/// </summary>
public class ConnectionTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Temp directory for all test files (auto-cleaned on disposal)
    /// </summary>
    public string TempDir => _tempDir;

    /// <summary>
    /// Initializes a new instance of the fixture.
    /// </summary>
    public ConnectionTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"ConnectionTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <inheritdoc/>
    public Task InitializeAsync()
    {
        // No async initialization needed
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
    /// Creates a unique empty test file for the calling test.
    /// Uses [CallerMemberName] to auto-populate the test name.
    /// </summary>
    /// <param name="testName">Auto-populated from caller method name</param>
    /// <returns>Path to the unique test file</returns>
    public string CreateTestFile([CallerMemberName] string testName = "")
    {
        var guid = Guid.NewGuid().ToString("N")[..8];
        var testFile = Path.Join(_tempDir, $"Conn_{testName}_{guid}.xlsx");
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(testFile, show: false);
        manager.CloseSession(sessionId, save: true);
        return testFile;
    }

    /// <summary>
    /// Creates a unique source file path (does not create the file).
    /// Used for OLEDB source workbooks that are created by test helpers.
    /// </summary>
    /// <param name="suffix">Optional suffix for the file name</param>
    /// <param name="testName">Auto-populated from caller method name</param>
    /// <returns>Path for the source file</returns>
    public string GetSourceFilePath(string suffix = "Source", [CallerMemberName] string testName = "")
    {
        return Path.Join(_tempDir, $"{testName}_{suffix}.xlsx");
    }
}




