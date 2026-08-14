using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Worksheet protection operations.
/// </summary>
public partial class SheetCommands
{
    /// <inheritdoc />
    public OperationResult SetProtection(IExcelBatch batch, string sheetName, bool isProtected, string? password = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                if (isProtected)
                {
                    if (string.IsNullOrWhiteSpace(password))
                    {
                        sheet.Protect();
                    }
                    else
                    {
                        sheet.Protect(password);
                    }
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(password))
                    {
                        sheet.Unprotect();
                    }
                    else
                    {
                        sheet.Unprotect(password);
                    }
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }

    /// <inheritdoc />
    public SheetProtectionResult GetProtection(IExcelBatch batch, string sheetName)
    {
        var result = new SheetProtectionResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName);
                if (sheet == null)
                {
                    throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                }

                result.IsProtected = sheet.ProtectContents;
                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref sheet);
            }
        });
    }
}
