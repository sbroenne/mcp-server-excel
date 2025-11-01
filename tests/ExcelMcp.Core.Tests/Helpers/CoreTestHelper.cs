using Sbroenne.ExcelMcp.Core.Commands;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Shared helper for Core integration tests.
/// Provides utilities for creating unique test files to ensure test isolation.
/// </summary>
public static class CoreTestHelper
{
    /// <summary>
    /// Creates a unique test file for isolated testing.
    /// Supports Excel files (.xlsx, .xlsm) and data files (.csv, .txt, etc.).
    /// </summary>
    /// <param name="testClassName">Name of the test class (e.g., "ParameterCommandsTests")</param>
    /// <param name="testName">Name of the test method (e.g., "Create_WithValidParameter_ReturnsSuccess")</param>
    /// <param name="tempDir">Temporary directory where the file will be created</param>
    /// <param name="extension">File extension including dot (e.g., ".xlsx", ".csv", ".txt")</param>
    /// <param name="content">Optional file content. Only used for non-Excel files. For CSV, defaults to sample data.</param>
    /// <returns>Absolute path to the created file</returns>
    /// <exception cref="InvalidOperationException">Thrown if Excel file creation fails</exception>
    /// <remarks>
    /// Usage patterns:
    /// <code>
    /// // Excel file
    /// var excelFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    ///     nameof(MyTests), nameof(MyTest), _tempDir, ".xlsx");
    ///
    /// // CSV file with default data
    /// var csvFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    ///     nameof(MyTests), nameof(MyTest), _tempDir, ".csv");
    ///
    /// // CSV file with custom data
    /// var csvFile = await CoreTestHelper.CreateUniqueTestFileAsync(
    ///     nameof(MyTests), nameof(MyTest), _tempDir, ".csv", "Col1,Col2\nA,B");
    /// </code>
    /// </remarks>
    public static async Task<string> CreateUniqueTestFileAsync(
        string testClassName,
        string testName,
        string tempDir,
        string extension = ".xlsx",
        string? content = null)
    {
        // Generate unique filename: ClassName_TestName_GUID.{extension}
        var fileName = $"{testClassName}_{testName}_{Guid.NewGuid():N}{extension}";
        var filePath = Path.Combine(tempDir, fileName);

        // Handle Excel files (.xlsx, .xlsm)
        if (extension == ".xlsx" || extension == ".xlsm")
        {
            var fileCommands = new FileCommands();
            var result = await fileCommands.CreateEmptyAsync(filePath, overwriteIfExists: false);

            if (!result.Success)
            {
                throw new InvalidOperationException(
                    $"Failed to create test Excel file '{filePath}': {result.ErrorMessage}. " +
                    "Excel may not be installed or the path may be invalid.");
            }
        }
        // Handle data files (.csv, .txt, etc.)
        else
        {
            if (content == null)
            {
                throw new ArgumentNullException(nameof(content),
                    $"Content must be provided for non-Excel files with extension '{extension}'");
            }

            File.WriteAllText(filePath, content);
        }

        return filePath;
    }
}
