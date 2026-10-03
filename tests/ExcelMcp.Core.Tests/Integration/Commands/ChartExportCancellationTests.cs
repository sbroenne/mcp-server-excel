using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Chart;
using Sbroenne.ExcelMcp.Core.Tests.Helpers;
using Sbroenne.ExcelMcp.Tests.Infrastructure;
using Xunit;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Tests.Integration.Commands;

[Collection("Sequential")]
[Trait("Category", "Integration")]
[Trait("Layer", "Core")]
[Trait("Feature", "ChartDepth")]
[Trait("RequiresExcel", "true")]
public sealed class ChartExportCancellationTests(ChartExportCancellationFixture fixture) :
    IClassFixture<ChartExportCancellationFixture>
{
    // Service cannot inject cancellation between batch admission and the command's STA callback.
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExportImage_CancelledCallbackDoesNotWriteOrReplaceOutput(bool existingOutput)
    {
        var path = fixture.CreateWorkbook();
        var output = Path.ChangeExtension(path, ".png");
        if (existingOutput)
            File.WriteAllText(output, "Keep existing image");
        using var inner = ExcelSession.BeginBatch(path);
        var sheetName = inner.Execute((context, _) =>
        {
            Excel.Sheets? sheets = null;
            Excel.Worksheet? sheet = null;
            Excel.Range? data = null;
            try
            {
                sheets = context.Book.Worksheets;
                sheet = (Excel.Worksheet)sheets[1];
                data = sheet.Range["A1:B4"];
                data.Value2 = new object[,] { { "Category", "Amount" }, { "A", 10 }, { "B", 20 }, { "C", 30 } };
                return sheet.Name;
            }
            finally
            {
                ComUtilities.Release(ref data);
                ComUtilities.Release(ref sheet);
                ComUtilities.Release(ref sheets);
            }
        });
        var commands = new ChartCommands();
        var created = commands.CreateFromRange(inner, sheetName, "A1:B4", ChartType.ColumnClustered, chartName: "Sales");
        Assert.True(created.Success, created.ErrorMessage);
        var before = inner.Execute((context, _) => (context.App.Calculation, context.App.ScreenUpdating));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        using var cancelled = new InjectedCancellationBatch(inner, cancellation.Token);

        var error = Assert.ThrowsAny<OperationCanceledException>(() =>
            commands.ExportImage(cancelled, "Sales", output, overwrite: existingOutput));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        if (existingOutput)
            Assert.Equal("Keep existing image", File.ReadAllText(output));
        else
            Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(fixture.DirectoryPath, $".{Path.GetFileNameWithoutExtension(output)}.*.tmp.png"));
        Assert.Equal(before, inner.Execute((context, _) => (context.App.Calculation, context.App.ScreenUpdating)));

        var followUp = commands.ExportImage(inner, "Sales", output, overwrite: existingOutput);
        Assert.True(followUp.Success, followUp.ErrorMessage);
        Assert.True(string.IsNullOrEmpty(followUp.ErrorMessage));
        Assert.Equal(output, followUp.FilePath);
        Assert.Equal((byte)137, File.ReadAllBytes(output)[0]);
    }
}

public sealed class ChartExportCancellationFixture : IAsyncLifetime
{
    public string DirectoryPath { get; } = Path.Combine(Path.GetTempPath(), $"chart-export-cancel-{Guid.NewGuid():N}");

    public Task InitializeAsync()
    {
        Directory.CreateDirectory(DirectoryPath);
        return Task.CompletedTask;
    }

    public string CreateWorkbook() =>
        SavedWorkbookTemplates.CopyBlankTo(Path.Combine(DirectoryPath, $"{Guid.NewGuid():N}.xlsx"));

    public Task DisposeAsync()
    {
        Directory.Delete(DirectoryPath, recursive: true);
        return Task.CompletedTask;
    }
}
