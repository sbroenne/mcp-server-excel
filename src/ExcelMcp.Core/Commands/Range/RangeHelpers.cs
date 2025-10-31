using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;


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

    /// <summary>
    /// Converts a value to a proper Excel cell value, handling System.Text.Json.JsonElement.
    /// MCP framework deserializes JSON arrays to JsonElement objects which cannot be marshalled to COM Variant.
    /// This helper detects JsonElement and converts to proper C# types before COM assignment.
    /// </summary>
    /// <param name="value">Value from MCP JSON deserialization or direct C# types</param>
    /// <returns>Proper C# type (string, long, double, bool) for COM marshalling</returns>
    public static object ConvertToCellValue(object? value)
    {
        if (value == null)
            return string.Empty;

        // Handle System.Text.Json.JsonElement (from MCP JSON deserialization)
        if (value is System.Text.Json.JsonElement jsonElement)
        {
            return jsonElement.ValueKind switch
            {
                System.Text.Json.JsonValueKind.String => jsonElement.GetString() ?? string.Empty,
                System.Text.Json.JsonValueKind.Number => jsonElement.TryGetInt64(out var i64) ? i64 : jsonElement.GetDouble(),
                System.Text.Json.JsonValueKind.True => true,
                System.Text.Json.JsonValueKind.False => false,
                System.Text.Json.JsonValueKind.Null => string.Empty,
                _ => jsonElement.ToString() ?? string.Empty
            };
        }

        // Already a proper type (from CLI or tests)
        return value;
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

        return await batch.Execute((ctx, ct) =>
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
