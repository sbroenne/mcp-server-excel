using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture for VBA tests providing efficient test file creation.
/// Each test gets a unique .xlsm file via CreateTestFile() method.
/// All files use .xlsm extension for VBA compatibility.
/// </summary>
public class VbaTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;

    /// <summary>
    /// Temp directory for all test files (auto-cleaned on disposal)
    /// </summary>
    public string TempDir => _tempDir;

    public VbaTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"VbaTests_{Guid.NewGuid():N}");
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
    /// Creates a unique empty .xlsm test file for the calling test.
    /// Uses [CallerMemberName] to auto-populate the test name.
    /// </summary>
    /// <param name="testName">Auto-populated from caller method name</param>
    /// <returns>Path to the unique .xlsm test file</returns>
    public string CreateTestFile([CallerMemberName] string testName = "")
    {
        var guid = Guid.NewGuid().ToString("N")[..8];
        var testFile = Path.Join(_tempDir, $"Vba_{testName}_{guid}.xlsm");
        using var manager = new SessionManager();
        var sessionId = manager.CreateSessionForNewFile(testFile, show: false);
        manager.CloseSession(sessionId, save: true);
        return testFile;
    }
}




