using Sbroenne.ExcelMcp.ComInterop.Session;

namespace Sbroenne.ExcelMcp.Core.Utilities;

internal static class WorkbookAccessGuard
{
    internal static void EnsureWritable(IExcelBatch batch, string? filePath = null)
    {
        batch.Execute((context, _) =>
        {
            var workbook = filePath is null ? context.Book : batch.GetWorkbook(filePath);
            if (workbook.ReadOnly)
            {
                throw new InvalidOperationException(
                    "Cannot change this workbook: Excel opened it read-only. " +
                    "This operation has not changed the workbook. Inspect its access and permissions before continuing.");
            }
        });
    }
}
