using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Service.Tests;

/// <summary>
/// Tests for FileCommands TestFile operation
/// </summary>
public sealed partial class ServiceFileCommandsTests
{
    [Theory]
    [InlineData("https://d.docs.live.net/123/Documents/Test.xlsx?download=1#sheet")]
    [InlineData("http://example.com:8080/Test.csv?download=1#sheet")]
    [InlineData("https://example.com/export.xlsx?report=one&token=test-first")]
    [InlineData("https://example.com/export.xlsx?report=two&token=test-second")]
    public async Task Open_RemoteLocation_ReachesExcelUnchangedWithoutRequiringVisibility(string location)
    {
        const string stoppedBeforeNetworkAccess = "Stopped at Excel opening boundary.";
        string? receivedLocation = null;
        IReadOnlyList<string>? reportedPaths = null;
        IReadOnlyList<string>? ownershipKeys = null;
        var previousHook = ExcelBatch.BeforeWorkbookOpenHook;
        using var service = new ExcelMcpService();
        try
        {
            ExcelBatch.BeforeWorkbookOpenHook = (path, _) =>
            {
                receivedLocation = path;
                var session = Assert.Single(service.SessionManager.ActiveSessionIds);
                reportedPaths = service.SessionManager.GetOwnedFilePaths(session);
                ownershipKeys = service.SessionManager.GetOwnedFilePathKeys(session);
                throw new InvalidOperationException(stoppedBeforeNetworkAccess);
            };
            var response = await service.ProcessAsync(new ServiceRequest
            {
                Command = "session.open",
                Args = System.Text.Json.JsonSerializer.Serialize(new { filePath = location, show = false })
            });

            Assert.Equal(location, receivedLocation);
            Assert.Equal(WorkbookLocation.Normalize(location), Assert.Single(Assert.IsAssignableFrom<IReadOnlyList<string>>(reportedPaths)));
            Assert.Equal(WorkbookLocation.GetOwnershipKey(location), Assert.Single(Assert.IsAssignableFrom<IReadOnlyList<string>>(ownershipKeys)));
            Assert.False(response.Success);
            Assert.Contains(stoppedBeforeNetworkAccess, response.ErrorMessage, StringComparison.Ordinal);
            Assert.Equal(0, service.SessionCount);
        }
        finally
        {
            ExcelBatch.BeforeWorkbookOpenHook = previousHook;
        }
    }

    [Fact]
    public void Test_ReadOnlyFile_CanBeOpenedWithoutChangingFile()
    {
        var path = _fixture.CreateTestFile();
        var attributes = File.GetAttributes(path);
        var bytes = File.ReadAllBytes(path);
        File.SetAttributes(path, attributes | FileAttributes.ReadOnly);
        try
        {
            var info = _fileCommands.Test(path);
            Assert.True(info.CanOpen, info.Message);
            Assert.True(info.IsValid, info.Message);
            Assert.Equal(bytes, File.ReadAllBytes(path));
        }
        finally
        {
            File.SetAttributes(path, attributes);
        }
    }

    [Fact]
    public void Test_ExistingValidFile_ReturnsSuccess()
    {
        // Arrange - Create a valid file
        var testFile = _fixture.CreateTestFile();
        var lastWriteTime = File.GetLastWriteTimeUtc(testFile);
        var bytes = File.ReadAllBytes(testFile);

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.True(info.IsValid);
        Assert.True(info.Success);
        Assert.True(info.CanOpen);
        Assert.False(info.IsIrmProtected);
        Assert.False(info.WillOpenReadOnly);
        Assert.False(info.RequiresVisibleSession);
        Assert.Equal(".xlsx", info.Extension);
        Assert.True(info.Size > 0);
        Assert.Null(info.Message);
        Assert.Equal(lastWriteTime, File.GetLastWriteTimeUtc(testFile));
        Assert.Equal(bytes, File.ReadAllBytes(testFile));
        Assert.Equal(0, _fixture.SessionCount);
    }
    [Fact]
    public void Test_NonExistent_ReturnsFailure()
    {
        // Arrange
        string testFile = Path.Join(_fixture.TempDir, $"NonExistent_{Guid.NewGuid():N}.xlsx");

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.False(info.Exists);
        Assert.False(info.IsValid);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.NotNull(info.Message);
        Assert.Contains("not found", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(testFile));
    }

    [Theory]
    [InlineData(".xlsx")]
    [InlineData(".xlsb")]
    public void Test_CorruptSupportedExtension_IsNotValidOrOpenable(string extension)
    {
        var testFile = Path.Join(_fixture.TempDir, $"Corrupt_{Guid.NewGuid():N}{extension}");
        System.IO.File.WriteAllText(testFile, "not an Excel workbook");
        var bytes = File.ReadAllBytes(testFile);

        var info = _fileCommands.Test(testFile);

        Assert.True(info.Exists);
        Assert.False(info.IsValid);
        Assert.False(info.Success);
        Assert.False(info.CanOpen);
        Assert.NotNull(info.Message);
        Assert.Contains("valid Excel workbook", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("already open", info.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(bytes, File.ReadAllBytes(testFile));
    }

    [Fact]
    public void Test_SavedBinaryWorkbook_IsValidWithoutChangingFile()
    {
        AssertSavedWorkbookIsValidWithoutChangingFile(".xlsb", Excel.XlFileFormat.xlExcel12);
    }

    [Fact]
    public void Test_SavedTemplate_IsValidWithoutChangingFile()
    {
        AssertSavedWorkbookIsValidWithoutChangingFile(".xltx", Excel.XlFileFormat.xlOpenXMLTemplate);
    }

    [Fact]
    [Trait("RunType", "OnDemand")]
    public void Test_SavedLegacyXlsWorkbook_IsValidWithoutChangingFile()
    {
        AssertSavedWorkbookIsValidWithoutChangingFile(".xls", Excel.XlFileFormat.xlExcel8);
    }

    private void AssertSavedWorkbookIsValidWithoutChangingFile(
        string extension, Excel.XlFileFormat format)
    {
        var source = _fixture.CreateTestFile();
        var saved = Path.ChangeExtension(source, extension);
        using (var batch = ExcelSession.BeginBatch(source))
        {
            batch.Execute((context, _) => context.Book.SaveAs(saved, format));
        }
        var bytes = File.ReadAllBytes(saved);
        var lastWriteTime = File.GetLastWriteTimeUtc(saved);

        var info = _fileCommands.Test(saved);

        Assert.True(info.Success, info.Message);
        Assert.True(info.Exists);
        Assert.True(info.IsValid);
        Assert.True(info.CanOpen);
        Assert.Equal(extension, info.Extension);
        Assert.Null(info.Message);
        Assert.Equal(bytes, File.ReadAllBytes(saved));
        Assert.Equal(lastWriteTime, File.GetLastWriteTimeUtc(saved));
        Assert.Equal(0, _fixture.SessionCount);
    }

    [Fact]
    public void Test_LockedSupportedFile_ReportsNotOpenable()
    {
        var testFile = _fixture.CreateTestFile();
        var bytes = File.ReadAllBytes(testFile);
        using (var lockStream = new FileStream(
            testFile,
            FileMode.Open,
            FileAccess.ReadWrite,
            FileShare.None))
        {
            var info = _fileCommands.Test(testFile);

            Assert.True(info.Exists);
            Assert.False(info.IsValid);
            Assert.False(info.Success);
            Assert.False(info.CanOpen);
            Assert.NotNull(info.Message);
            Assert.Contains("Excel could not open", info.Message, StringComparison.OrdinalIgnoreCase);
        }
        Assert.Equal(bytes, File.ReadAllBytes(testFile));
        Assert.True(_fileCommands.Test(testFile).CanOpen);
    }

    [Theory]
    [InlineData("TestFile.csv", ".csv")]
    [InlineData("TestFile.txt", ".txt")]
    public void Test_TextFile_UsesExcelToDetermineOpenability(string fileName, string expectedExt)
    {
        // Arrange
        string testFile = Path.Join(_fixture.TempDir, $"{Guid.NewGuid():N}_{fileName}");

        System.IO.File.WriteAllText(testFile, "native-open-marker");

        // Act
        var info = _fileCommands.Test(testFile);

        // Assert
        Assert.True(info.Exists);
        Assert.True(info.IsValid, info.Message);
        Assert.True(info.Success, info.Message);
        Assert.True(info.CanOpen, info.Message);
        Assert.Equal(expectedExt, info.Extension);
        Assert.Null(info.Message);
        Assert.Equal("native-open-marker", File.ReadAllText(testFile));
    }
}
