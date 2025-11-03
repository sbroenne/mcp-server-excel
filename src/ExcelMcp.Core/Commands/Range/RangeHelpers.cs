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
    /// Returns null if resolution fails.
    /// </summary>
    public static dynamic? ResolveRange(dynamic book, string sheetName, string rangeAddress, out string? specificError)
    {
        specificError = null;

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
                // List available named ranges for helpful error
                List<string> availableRanges = new();
                dynamic? names = null;
                try
                {
                    names = book.Names;
                    for (int i = 1; i <= Math.Min(names.Count, 10); i++)
                    {
                        dynamic? name = null;
                        try
                        {
                            name = names.Item(i);
                            availableRanges.Add(name.Name);
                        }
                        finally
                        {
                            ComUtilities.Release(ref name);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref names);
                }

                if (availableRanges.Count > 0)
                {
                    string rangeList = string.Join(", ", availableRanges);
                    if (names?.Count > 10)
                    {
                        rangeList += $" ... ({names.Count - 10} more)";
                    }
                    specificError = $"Named range '{rangeAddress}' not found. Available named ranges: {rangeList}";
                }
                else
                {
                    specificError = $"Named range '{rangeAddress}' not found. No named ranges exist in this workbook.";
                }
                
                specificError += " Use excel_namedrange(action: 'list') to see all, or excel_namedrange(action: 'create') to create one.";
                return null;
            }
        }

        // Regular range (sheet + address)
        // First check if sheet exists
        dynamic? sheet = null;
        try
        {
            sheet = ComUtilities.FindSheet(book, sheetName);
            if (sheet == null)
            {
                // List available sheets for helpful error
                List<string> availableSheets = new();
                dynamic? sheets = null;
                try
                {
                    sheets = book.Worksheets;
                    for (int i = 1; i <= Math.Min(sheets.Count, 10); i++)
                    {
                        dynamic? ws = null;
                        try
                        {
                            ws = sheets.Item(i);
                            availableSheets.Add(ws.Name);
                        }
                        finally
                        {
                            ComUtilities.Release(ref ws);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref sheets);
                }

                if (availableSheets.Count > 0)
                {
                    string sheetList = string.Join(", ", availableSheets);
                    if (sheets?.Count > 10)
                    {
                        sheetList += $" ... ({sheets.Count - 10} more)";
                    }
                    specificError = $"Sheet '{sheetName}' not found. Available sheets: {sheetList}";
                }
                else
                {
                    specificError = $"Sheet '{sheetName}' not found. Workbook has no worksheets.";
                }
                
                specificError += " Use excel_worksheet(action: 'list') to see all sheets.";
                return null;
            }

            // Sheet exists, now try to get the range
            try
            {
                return sheet.Range[rangeAddress];
            }
            catch (Exception ex)
            {
                specificError = $"Sheet '{sheetName}' exists, but range '{rangeAddress}' is invalid. " +
                               $"Error: {ex.Message}. " +
                               $"Verify the range address format (e.g., 'A1:E10', 'A1', 'A:A').";
                return null;
            }
        }
        finally
        {
            ComUtilities.Release(ref sheet);
        }
    }

    /// <summary>
    /// Resolves a range address to a Range COM object (backward compatibility).
    /// Supports both regular ranges (Sheet1!A1:D10) and named ranges.
    /// </summary>
    public static dynamic? ResolveRange(dynamic book, string sheetName, string rangeAddress)
    {
        string? ignoredError;
        return ResolveRange(book, sheetName, rangeAddress, out ignoredError);
    }

    /// <summary>
    /// Gets appropriate error message for range resolution failure
    /// </summary>
    public static string GetResolveError(string sheetName, string rangeAddress)
    {
        if (string.IsNullOrEmpty(sheetName))
        {
            return $"Named range '{rangeAddress}' not found. " +
                   $"Use excel_namedrange(action: 'list') to see available named ranges, " +
                   $"or create it with excel_namedrange(action: 'create').";
        }
        return $"Sheet '{sheetName}' or range '{rangeAddress}' not found. " +
               $"Use excel_worksheet(action: 'list') to see available sheets, " +
               $"or verify the range address is correct (e.g., 'A1:E10').";
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
