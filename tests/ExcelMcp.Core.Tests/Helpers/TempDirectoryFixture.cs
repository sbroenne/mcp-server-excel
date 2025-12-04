using System.Runtime.CompilerServices;

using Sbroenne.ExcelMcp.Core.Commands;

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
///     private readonly TempDirectoryFixture _fixture;
///     
///     public MyTests(TempDirectoryFixture fixture)
///     {
///         _fixture = fixture;
///     }
/// }
/// </code>
/// </remarks>
public class TempDirectoryFixture : IDisposable
{
    private readonly FileCommands _fileCommands = new();

    /// <summary>
    /// Temporary directory for test files. Created in constructor, deleted in Dispose.
    /// Shared across all tests in the test class.
    /// </summary>
    public string TempDir { get; }

    /// <summary>
    /// Creates a unique test Excel file for the calling test method.
    /// </summary>
    /// <param name="testName">Auto-populated with the calling method name.</param>
    /// <param name="extension">File extension (default: .xlsx).</param>
    /// <returns>Full path to the created file.</returns>
    public string CreateTestFile([CallerMemberName] string testName = "", string extension = ".xlsx")
    {
        var fileName = $"{testName}_{Guid.NewGuid():N}{extension}";
        var filePath = Path.Combine(TempDir, fileName);
        _fileCommands.CreateEmpty(filePath);
        return filePath;
    }

    private bool _disposed;

    /// <summary>
    /// Initializes the fixture with a unique temp directory.
    /// </summary>
    public TempDirectoryFixture()
    {
        TempDir = Path.Join(Path.GetTempPath(), $"ExcelMcp_Tests_{Guid.NewGuid():N}");
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
