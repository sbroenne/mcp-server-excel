using Xunit;

namespace Sbroenne.ExcelMcp.Core.Tests.Commands.File;

/// <summary>
/// Tests for rapid sequential CreateEmpty operations - simulates LLM test scenario
/// where multiple files are created in quick succession.
///
/// This test catches COM apartment deadlocks that occur when Task.Run() is used
/// inside STA threads for COM operations.
/// </summary>
public partial class FileCommandsTests
{
    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Speed", "Slow")]
    [Trait("Layer", "Core")]
    [Trait("Feature", "Files")]
    [Trait("RequiresExcel", "true")]
    public void CreateEmpty_RapidSequentialCreation_AllFilesCreated()
    {
        // Arrange - Create multiple files in rapid succession
        // This simulates the LLM test scenario where each session creates a new file
        const int fileCount = 5;
        var testFiles = new string[fileCount];

        for (int i = 0; i < fileCount; i++)
        {
            testFiles[i] = Path.Join(
                _fixture.TempDir,
                $"{nameof(CreateEmpty_RapidSequentialCreation_AllFilesCreated)}_{i}_{Guid.NewGuid():N}.xlsx");
        }

        // Act - Create files in rapid succession (no delay between calls)
        for (int i = 0; i < fileCount; i++)
        {
            _fileCommands.CreateEmpty(testFiles[i]);
        }

        // Assert - All files were created successfully
        for (int i = 0; i < fileCount; i++)
        {
            Assert.True(
                System.IO.File.Exists(testFiles[i]),
                $"File {i} should exist after CreateEmpty: {testFiles[i]}");

            var fileInfo = new FileInfo(testFiles[i]);
            Assert.True(
                fileInfo.Length > 0,
                $"File {i} should have content (not empty): {testFiles[i]}");
        }
    }

    [Fact]
    [Trait("Category", "Integration")]
    [Trait("Speed", "Slow")]
    [Trait("Layer", "Core")]
    [Trait("Feature", "Files")]
    [Trait("RequiresExcel", "true")]
    public void CreateEmpty_RapidSequentialWithMixedExtensions_AllFilesCreated()
    {
        // Arrange - Mix of .xlsx and .xlsm files
        var testFiles = new[]
        {
            Path.Join(_fixture.TempDir, $"rapid_test_1_{Guid.NewGuid():N}.xlsx"),
            Path.Join(_fixture.TempDir, $"rapid_test_2_{Guid.NewGuid():N}.xlsm"),
            Path.Join(_fixture.TempDir, $"rapid_test_3_{Guid.NewGuid():N}.xlsx"),
            Path.Join(_fixture.TempDir, $"rapid_test_4_{Guid.NewGuid():N}.xlsm"),
            Path.Join(_fixture.TempDir, $"rapid_test_5_{Guid.NewGuid():N}.xlsx"),
        };

        // Act - Create files in rapid succession
        foreach (var file in testFiles)
        {
            _fileCommands.CreateEmpty(file);
        }

        // Assert - All files were created
        foreach (var file in testFiles)
        {
            Assert.True(System.IO.File.Exists(file), $"File should exist: {file}");
        }
    }
}
