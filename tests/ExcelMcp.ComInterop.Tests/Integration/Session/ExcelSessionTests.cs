using Sbroenne.ExcelMcp.ComInterop.Session;
using Xunit;
using Xunit.Abstractions;

namespace Sbroenne.ExcelMcp.ComInterop.Tests.Integration;

/// <summary>
/// Integration tests for ExcelSession.BeginBatch input validation.
/// ExcelBatchTests covers native batch creation and owned-process cleanup.
/// </summary>
[Trait("Category", "Integration")]
[Trait("Speed", "Slow")]
[Trait("Layer", "ComInterop")]
[Trait("Feature", "ExcelSession")]
[Collection("Sequential")] // Disable parallelization to avoid COM interference
[Trait("RequiresExcel", "true")]
public class ExcelSessionTests : IDisposable
{
    private readonly ITestOutputHelper _output;

    public ExcelSessionTests(ITestOutputHelper output)
    {
        _output = output;

    }

    /// <summary>
    /// Runs after each test
    /// </summary>
    public void Dispose()
    {
        // Nothing to dispose
        GC.SuppressFinalize(this);
    }

    [Fact]
    public void BeginBatch_WithNonExistentFile_ThrowsFileNotFoundException()
    {
        // Arrange
        string nonExistentFile = Path.Join(Path.GetTempPath(), $"does-not-exist-{Guid.NewGuid():N}.xlsx");

        // Act & Assert
        Assert.Throws<FileNotFoundException>(() =>
        {
            using var batch = ExcelSession.BeginBatch(nonExistentFile);
        });
        Assert.False(File.Exists(nonExistentFile));

        _output.WriteLine("✓ Correctly throws FileNotFoundException for non-existent file");
    }

    [Fact]
    public void BeginBatch_WithInvalidExtension_ThrowsArgumentException()
    {
        // Arrange
        string invalidFile = Path.Join(Path.GetTempPath(), $"test-{Guid.NewGuid():N}.txt");
        File.WriteAllText(invalidFile, "dummy");

        try
        {
            // Act & Assert
            var exception = Assert.Throws<ArgumentException>(() =>
            {
                using var batch = ExcelSession.BeginBatch(invalidFile);
            });

            Assert.Contains("Invalid file extension", exception.Message);
            Assert.Equal("dummy", File.ReadAllText(invalidFile));
            _output.WriteLine("✓ Correctly rejects non-Excel file extension");
        }
        finally
        {
            if (File.Exists(invalidFile)) File.Delete(invalidFile);
        }
    }
}
