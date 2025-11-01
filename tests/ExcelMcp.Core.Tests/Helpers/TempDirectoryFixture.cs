namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// xUnit test fixture that provides temp directory management for integration tests.
/// Automatically creates a unique temp directory and cleans up on disposal.
/// </summary>
/// <remarks>
/// Usage with xUnit IClassFixture:
/// <code>
/// public partial class MyTests : IClassFixture&lt;TempDirectoryFixture&gt;
/// {
///     private readonly string _tempDir;
///     
///     public MyTests(TempDirectoryFixture fixture)
///     {
///         _tempDir = fixture.TempDir;
///     }
/// }
/// </code>
/// </remarks>
public class TempDirectoryFixture : IDisposable
{
    /// <summary>
    /// Temporary directory for test files. Created in constructor, deleted in Dispose.
    /// Shared across all tests in the test class.
    /// </summary>
    public string TempDir { get; }

    private bool _disposed;

    /// <summary>
    /// Initializes the fixture with a unique temp directory.
    /// </summary>
    public TempDirectoryFixture()
    {
        TempDir = Path.Combine(Path.GetTempPath(), $"ExcelMcp_Tests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(TempDir);
    }

    /// <summary>
    /// Cleans up the temporary directory and all files within it.
    /// </summary>
    public void Dispose()
    {
        if (_disposed) return;

        try
        {
            if (Directory.Exists(TempDir))
            {
                Directory.Delete(TempDir, recursive: true);
            }
        }
        catch
        {
            // Cleanup failures are non-critical and shouldn't fail tests
        }

        _disposed = true;
        GC.SuppressFinalize(this);
    }
}
