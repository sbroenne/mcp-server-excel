using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Xunit;

namespace Sbroenne.ExcelMcp.CLI.Tests.Integration;

/// <summary>
/// Base class for CLI daemon integration tests.
/// Provides temp directory management and session lifecycle helpers.
/// Tests daemon handlers by calling Core Commands directly (same as daemon does).
/// </summary>
public abstract class DaemonIntegrationTestBase : IClassFixture<TempDirectoryFixture>, IDisposable
{
    protected TempDirectoryFixture Fixture { get; }
    protected string TestFilePath { get; }
    private bool _disposed;

    protected DaemonIntegrationTestBase(TempDirectoryFixture fixture)
    {
        Fixture = fixture;
        TestFilePath = CoreTestHelper.CreateUniqueTestFile(
            GetType().Name,
            "SharedTest",
            fixture.TempDir);
    }

    /// <summary>
    /// Creates a unique test file for isolated testing.
    /// </summary>
    protected string CreateTestFile(string testName, string extension = ".xlsx")
    {
        return CoreTestHelper.CreateUniqueTestFile(
            GetType().Name,
            testName,
            Fixture.TempDir,
            extension);
    }

    /// <summary>
    /// Creates a batch session for the shared test file.
    /// </summary>
    protected IExcelBatch CreateBatch() => ExcelSession.BeginBatch(TestFilePath);

    /// <summary>
    /// Creates a batch session for a specific file.
    /// </summary>
    protected static IExcelBatch CreateBatch(string filePath) => ExcelSession.BeginBatch(filePath);

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
