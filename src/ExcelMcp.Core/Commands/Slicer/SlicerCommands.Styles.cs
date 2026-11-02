using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Workbook;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Slicer;

public sealed partial class SlicerCommands
{
    /// <inheritdoc />
    public TableStyleListResult ListStyles(IExcelBatch batch)
    {
        var result = new WorkbookCommands().ListTableStyles(batch);
        result.Styles.RemoveAll(style => !style.ShowAsAvailableSlicerStyle);
        return result;
    }

    /// <inheritdoc />
    public TableStyleResult CreateStyle(IExcelBatch batch, string styleName, string sourceStyleName,
        SlicerStyleOptions slicerStyleOptions)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        ArgumentException.ThrowIfNullOrWhiteSpace(sourceStyleName);
        ArgumentNullException.ThrowIfNull(slicerStyleOptions);
        var options = slicerStyleOptions.ToTableStyleOptions();
        WorkbookCommands.ValidateTableStyleOptions(options);
        return batch.Execute((ctx, ct) =>
        {
            Excel.TableStyles? styles = null;
            Excel.TableStyle? existing = null;
            Excel.TableStyle? source = null;
            Excel.TableStyle? style = null;
            try
            {
                styles = ctx.Book.TableStyles;
                existing = WorkbookCommands.TryFindTableStyle(styles, styleName, ct);
                if (existing is not null)
                    throw new ArgumentException($"Style '{styleName}' already exists.", nameof(styleName));
                source = WorkbookCommands.FindTableStyle(styles, sourceStyleName, ct);
                if (!source.ShowAsAvailableSlicerStyle)
                    throw new ArgumentException($"Style '{sourceStyleName}' is not available for ordinary slicers.", nameof(sourceStyleName));
                ct.ThrowIfCancellationRequested();
                style = source.Duplicate(styleName);
                try
                {
                    ComUtilities.Release(ref style);
                    style = styles[styleName];
                    WorkbookCommands.ApplyTableStyleOptions(style, options, ct);
                    return WorkbookCommands.ReadTableStyleResult(style, batch.WorkbookPath, ct);
                }
                catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
                {
                    throw CreatedObjectFailure.CreateDescribed("Configuring the slicer style",
                        $"style '{styleName}'",
                        $"Style '{styleName}' remains in the workbook and may be partially formatted. Inspect it with workbook get-table-style before deciding what to change.",
                        ex);
                }
            }
            finally
            {
                ComUtilities.Release(ref style);
                ComUtilities.Release(ref source);
                ComUtilities.Release(ref existing);
                ComUtilities.Release(ref styles);
            }
        });
    }

    /// <inheritdoc />
    public SlicerStateResult ApplyStyle(IExcelBatch batch, string slicerName, string styleName)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(slicerName);
        ArgumentException.ThrowIfNullOrWhiteSpace(styleName);
        return UpdateSlicerCore(batch, slicerName, new SlicerUpdateOptions { Style = styleName }, ordinaryOnly: true);
    }
}
