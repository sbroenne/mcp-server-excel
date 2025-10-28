using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

namespace Sbroenne.ExcelMcp.Core.Commands.Range;

/// <summary>
/// Range resolution and helper methods for RangeCommands
/// </summary>
public static class RangeHelpers
{
    /// <summary>
    /// Resolves a range address to a Range COM object.
    /// Supports both regular ranges (Sheet1!A1:D10) and named ranges.
    /// </summary>
    public static dynamic? ResolveRange(dynamic book, string sheetName, string rangeAddress)
    {
        // Named range (empty sheetName)
        if (string.IsNullOrEmpty(sheetName))
        {
            try
            {
                dynamic names = book.Names;
                dynamic name = names.Item(rangeAddress);
                return name.RefersToRange;
            }
            catch
            {
                return null;
            }
        }

        // Regular range (sheet + address)
        dynamic? sheet = null;
        try
        {
            sheet = ComUtilities.FindSheet(book, sheetName);
            if (sheet == null) return null;

            return sheet.Range[rangeAddress];
        }
        finally
        {
            ComUtilities.Release(ref sheet);
        }
    }

    /// <summary>
    /// Gets appropriate error message for range resolution failure
    /// </summary>
    public static string GetResolveError(string sheetName, string rangeAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            return $"Named range '{rangeAddress}' not found";
        }
        return $"Sheet '{sheetName}' or range '{rangeAddress}' not found";
    }
}

/// <summary>
/// Internal helper methods for RangeCommands partial class
/// </summary>
public partial class RangeCommands
{
    /// <summary>
    /// Helper for clear operations
    /// </summary>
    private async Task<OperationResult> ClearRangeAsync(
        IExcelBatch batch,
        string sheetName,
        string rangeAddress,
        string action,
        Action<dynamic> clearAction)
    {
        var result = new OperationResult { FilePath = batch.WorkbookPath, Action = action };

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? range = null;
            try
            {
                range = RangeHelpers.ResolveRange(ctx.Book, sheetName, rangeAddress);
                if (range == null)
                {
                    result.Success = false;
                    result.ErrorMessage = RangeHelpers.GetResolveError(sheetName, rangeAddress);
                    return result;
                }

                clearAction(range);
                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = ex.Message;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref range);
            }
        });
    }
}
