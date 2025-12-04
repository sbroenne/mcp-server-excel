using System.Diagnostics;
using System.Runtime.CompilerServices;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands;
using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Helpers;

/// <summary>
/// Fixture that creates ONE Range test file per test CLASS.
/// Each test uses different sheets within the same file for isolation.
/// - Created ONCE before any tests run (~3-5s)
/// - Shared by all tests in the class
/// - Each test gets its own batch AND its own sheet (isolation)
/// - Reduces file creation overhead by ~95%
/// </summary>
public class RangeTestsFixture : IAsyncLifetime
{
    private readonly string _tempDir;
    private readonly FileCommands _fileCommands = new();
    private int _sheetCounter;

    /// <summary>
    /// Temp directory for all test files (auto-cleaned on disposal)
    /// </summary>
    public string TempDir => _tempDir;

    /// <summary>
    /// Path to the test Range file
    /// </summary>
    public string TestFilePath { get; private set; } = null!;

    /// <summary>
    /// Results of fixture creation (exposed for validation)
    /// </summary>
    public FixtureCreationResult CreationResult { get; private set; } = null!;

    /// <summary>
    /// Initializes a new instance of the fixture
    /// </summary>
    public RangeTestsFixture()
    {
        _tempDir = Path.Join(Path.GetTempPath(), $"RangeTests_{Guid.NewGuid():N}");
        Directory.CreateDirectory(_tempDir);
    }

    /// <summary>
    /// Gets a unique sheet name for test isolation.
    /// Each test should call this to get its own sheet.
    /// Sheet names are limited to 31 chars (Excel limit).
    /// </summary>
    public string GetUniqueSheetName([CallerMemberName] string testName = "")
    {
        var counter = Interlocked.Increment(ref _sheetCounter);
        // Limit sheet name to 31 chars (Excel limit) - format: "T001_TestMethodName"
        var prefix = $"T{counter:D3}_";
        var maxNameLength = 31 - prefix.Length;
        var shortName = testName.Length > maxNameLength ? testName[..maxNameLength] : testName;
        return $"{prefix}{shortName}";
    }

    /// <summary>
    /// Creates a unique sheet for a test and returns its name.
    /// Call this in the Arrange phase of each test.
    /// </summary>
    public string CreateTestSheet(IExcelBatch batch, [CallerMemberName] string testName = "")
    {
        var sheetName = GetUniqueSheetName(testName);

        batch.Execute((ctx, ct) =>
        {
            dynamic sheets = ctx.Book.Worksheets;
            dynamic newSheet = sheets.Add(After: sheets.Item(sheets.Count));
            newSheet.Name = sheetName;
            return 0;
        });

        return sheetName;
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
        _fileCommands.CreateEmpty(filePath);
        return filePath;
    }

    /// <summary>
    /// Called ONCE before any tests in the class run.
    /// Creates a workbook that tests will add sheets to.
    /// </summary>
    public Task InitializeAsync()
    {
        var sw = Stopwatch.StartNew();

        TestFilePath = Path.Join(_tempDir, "RangeTest.xlsx");
        CreationResult = new FixtureCreationResult();

        try
        {
            var fileCommands = new FileCommands();
            fileCommands.CreateEmpty(TestFilePath);
            CreationResult.FileCreated = true;

            sw.Stop();
            CreationResult.Success = true;
            CreationResult.CreationTimeSeconds = sw.Elapsed.TotalSeconds;
        }
        catch (Exception ex)
        {
            CreationResult.Success = false;
            CreationResult.ErrorMessage = ex.Message;
            sw.Stop();
            throw;
        }

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

/// <summary>
/// Generic fixture creation result
/// </summary>
public class FixtureCreationResult
{
    /// <summary>Whether fixture creation succeeded</summary>
    public bool Success { get; set; }

    /// <summary>Whether the Excel file was created</summary>
    public bool FileCreated { get; set; }

    /// <summary>Time taken to create the fixture</summary>
    public double CreationTimeSeconds { get; set; }

    /// <summary>Error message if creation failed</summary>
    public string? ErrorMessage { get; set; }
}
