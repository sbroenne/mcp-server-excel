using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Chart;

public partial class ChartCommands
{
    /// <inheritdoc />
    public OperationResult ExportImage(
        IExcelBatch batch, string chartName, string targetPath, ChartImageFormat imageFormat = ChartImageFormat.Png, bool overwrite = false)
    {
        if (!Enum.IsDefined(imageFormat)) throw new ArgumentOutOfRangeException(nameof(imageFormat));
        var output = WorkbookCommands.ValidateOutputPath(targetPath, overwrite);
        var extension = Path.GetExtension(output).ToLowerInvariant();
        var matches = imageFormat switch
        {
            ChartImageFormat.Png => extension == ".png",
            ChartImageFormat.Jpeg => extension is ".jpg" or ".jpeg",
            ChartImageFormat.Gif => extension == ".gif",
            _ => false
        };
        if (!matches) throw new ArgumentException("Output extension must match Png (.png), Jpeg (.jpg/.jpeg) or Gif (.gif).", nameof(targetPath));
        var writePath = WorkbookCommands.GetWritePath(output, overwrite);
        try
        {
            var result = batch.Execute((ctx, ct) =>
            {
                var found = FindChart(ctx.Book, chartName);
                Excel.Chart? chart = null;
                try
                {
                    chart = (Excel.Chart?)found.Chart ?? throw new ArgumentException($"Chart '{chartName}' not found.");
                    found.Chart = null;
                    ct.ThrowIfCancellationRequested();
                    var filter = imageFormat == ChartImageFormat.Jpeg ? "JPG" : imageFormat.ToString().ToUpperInvariant();
                    if (!chart.Export(writePath, filter, false) || !File.Exists(writePath) || new FileInfo(writePath).Length == 0)
                        throw new IOException($"Excel failed to export a nonempty {imageFormat} image. The installed Excel image filter may be unavailable.");
                    return new OperationResult { Success = true, FilePath = output, Action = "export-image" };
                }
                finally
                {
                    ComUtilities.Release(ref chart);
                    if (found.Shape != null) ComUtilities.Release(ref found.Shape!);
                    if (found.Chart != null) ComUtilities.Release(ref found.Chart!);
                }
            });
            WorkbookCommands.CommitOutput(writePath, output);
            return result;
        }
        catch (Exception exportError)
        {
            try
            {
                File.Delete(writePath);
            }
            catch (Exception cleanupError)
            {
                throw new AggregateException("Chart image export and removal of its failed output both failed.", exportError, cleanupError);
            }
            throw;
        }
    }
}
