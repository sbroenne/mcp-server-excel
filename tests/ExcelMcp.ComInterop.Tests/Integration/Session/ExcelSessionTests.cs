using Sbroenne.ExcelMcp.ComInterop.Session;
using System.Runtime.InteropServices;
using Xunit;
using Xunit.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

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
    public void BeginBatch_WithNonExistentFile_PreservesExcelOpenFailure()
    {
        // Arrange
        string nonExistentFile = Path.Join(Path.GetTempPath(), $"does-not-exist-{Guid.NewGuid():N}.xlsx");

        // Act & Assert
        var exception = Assert.Throws<COMException>(() =>
        {
            using var batch = ExcelSession.BeginBatch(nonExistentFile);
        });
        Assert.False(File.Exists(nonExistentFile));
        Assert.Contains(Path.GetFileName(nonExistentFile), exception.Message, StringComparison.OrdinalIgnoreCase);
        _output.WriteLine($"Excel rejected the missing file: 0x{exception.HResult:X8}");
    }

    [Theory]
    [InlineData(".txt")]
    [InlineData(".csv")]
    public void BeginBatch_WithTextFile_OpensItsContentsInExcel(string extension)
    {
        // Arrange
        string textFile = Path.Join(Path.GetTempPath(), $"test-{Guid.NewGuid():N}{extension}");
        File.WriteAllText(textFile, "native-open-marker");

        try
        {
            using (var batch = ExcelSession.BeginBatch(textFile))
            {
                batch.Execute((context, _) =>
                {
                    Excel.Sheets? sheets = null;
                    Excel.Worksheet? sheet = null;
                    Excel.Range? cell = null;
                    try
                    {
                        sheets = context.Book.Worksheets;
                        sheet = (Excel.Worksheet)sheets[1];
                        cell = sheet.Range["A1"];
                        Assert.Equal("native-open-marker", (string)cell.Value2);
                        Assert.False(context.Book.ReadOnly);
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                        ComUtilities.Release(ref sheet);
                        ComUtilities.Release(ref sheets);
                    }
                });
                Assert.IsType<ExcelBatch>(batch).Close();
            }
            Assert.Equal("native-open-marker", File.ReadAllText(textFile));
        }
        finally
        {
            if (File.Exists(textFile)) File.Delete(textFile);
        }
    }
}
