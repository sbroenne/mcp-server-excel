using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Shared helper for Core integration tests.
/// Provides utilities for creating unique test files to ensure test isolation.
/// </summary>
public static class CoreTestHelper
{
    /// <summary>
    /// Creates a unique test Excel file for isolated testing.
    /// Each test should call this method to get its own unique file, preventing test pollution.
    /// </summary>
    /// <param name="testClassName">Name of the test class (e.g., "ParameterCommandsTests")</param>
    /// <param name="testName">Name of the test method (e.g., "Create_WithValidParameter_ReturnsSuccess")</param>
    /// <param name="tempDir">Temporary directory where the file will be created</param>
    /// <returns>Absolute path to the created Excel file</returns>
    /// <exception cref="InvalidOperationException">Thrown if file creation fails</exception>
    /// <remarks>
    /// Usage pattern:
    /// <code>
    /// private readonly string _tempDir;
    /// 
    /// public MyTests()
    /// {
    ///     _tempDir = Path.Combine(Path.GetTempPath(), $"MyTests_{Guid.NewGuid():N}");
    ///     Directory.CreateDirectory(_tempDir);
    /// }
    /// 
    /// [Fact]
    /// public async Task MyTest()
    /// {
    ///     var testFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    ///         nameof(MyTests), nameof(MyTest), _tempDir);
    ///     await using var batch = await ExcelSession.BeginBatchAsync(testFile);
    ///     // ... test operations ...
    /// }
    /// </code>
    /// </remarks>
    public static async Task<string> CreateUniqueTestFileAsync(
        string testClassName,
        string testName,
        string tempDir)
    {
        // Generate unique filename: ClassName_TestName_GUID.xlsx
        var fileName = $"{testClassName}_{testName}_{Guid.NewGuid():N}.xlsx";
        var filePath = Path.Combine(tempDir, fileName);

        var fileCommands = new FileCommands();
        var result = await fileCommands.CreateEmptyAsync(filePath, overwriteIfExists: false);

        if (!result.Success)
        {
            throw new InvalidOperationException(
                $"Failed to create test Excel file '{filePath}': {result.ErrorMessage}. " +
                "Excel may not be installed or the path may be invalid.");
        }

        return filePath;
    }
}
