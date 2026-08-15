using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// Workbook-level command implementation.
/// </summary>
public partial class WorkbookCommands : IWorkbookCommands
{
    /// <inheritdoc />
    public OperationResult SetProtection(IExcelBatch batch, bool isProtected, string? password = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            if (isProtected)
            {
                ctx.Book.Protect(password, Structure: true, Windows: false);
            }
            else if (string.IsNullOrWhiteSpace(password))
            {
                ctx.Book.Unprotect();
            }
            else
            {
                ctx.Book.Unprotect(password);
            }

            return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
        });
    }

    /// <inheritdoc />
    public WorkbookProtectionResult GetProtection(IExcelBatch batch)
    {
        var result = new WorkbookProtectionResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            result.IsProtected = ctx.Book.ProtectStructure || ctx.Book.ProtectWindows;
            result.Success = true;
            return result;
        });
    }

    /// <inheritdoc />
    public OperationResult SetViewOptions(IExcelBatch batch, bool? displayGridlines = null, bool? displayHeadings = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                windows = ctx.Book.Windows;
                window = windows[1];
                if (displayGridlines.HasValue)
                {
                    window.DisplayGridlines = displayGridlines.Value;
                }

                if (displayHeadings.HasValue)
                {
                    window.DisplayHeadings = displayHeadings.Value;
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        });
    }

    /// <inheritdoc />
    public WorkbookViewOptionsResult GetViewOptions(IExcelBatch batch)
    {
        var result = new WorkbookViewOptionsResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Windows? windows = null;
            Excel.Window? window = null;
            try
            {
                windows = ctx.Book.Windows;
                window = windows[1];
                result.DisplayGridlines = window.DisplayGridlines;
                result.DisplayHeadings = window.DisplayHeadings;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref window);
                ComUtilities.Release(ref windows);
            }
        });
    }
}
