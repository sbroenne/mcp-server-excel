using System.Runtime.InteropServices;
using System.Text.Json;
using Sbroenne.ExcelMcp.Core.Commands;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Utilities;

/// <summary>
/// Shared exception classification for Service and local adapter failures.
/// Classification never depends on arbitrary exception text.
/// </summary>
public static class OperationFailureClassifier
{
    /// <summary>
    /// Finds the COM error code through context wrappers without choosing an
    /// arbitrary cause from a multi-failure aggregate.
    /// </summary>
    public static string? GetComHResult(Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);
        for (Exception? current = exception; current != null; current = current.InnerException)
        {
            if (current is COMException com)
            {
                return $"0x{com.HResult:X8}";
            }
            if (current is AggregateException aggregate && aggregate.InnerExceptions.Count != 1)
            {
                return null;
            }
        }
        return null;
    }

    /// <summary>
    /// Preserves the outermost known meaning, looking through context-only wrappers.
    /// Mixed or partly unknown aggregate failures remain unclassified.
    /// </summary>
    public static string? Classify(Exception exception)
    {
        ArgumentNullException.ThrowIfNull(exception);

        for (Exception? current = exception; current != null; current = current.InnerException)
        {
            var category = current switch
            {
                OperationFailureException failure => failure.ErrorCategory.ToString(),
                PowerQueryCommandException query => query.ErrorCategory,
                TimeoutException => "Timeout",
                OperationCanceledException => "Cancelled",
                ArgumentException or JsonException => "InvalidInput",
                COMException => "ComInterop",
                _ => null
            };
            if (category != null)
            {
                return category;
            }

            if (current is AggregateException aggregate)
            {
                string? sharedCategory = null;
                foreach (var inner in aggregate.InnerExceptions)
                {
                    var innerCategory = Classify(inner);
                    if (innerCategory == null || (sharedCategory != null && sharedCategory != innerCategory))
                    {
                        return null;
                    }
                    sharedCategory = innerCategory;
                }
                return sharedCategory;
            }
        }

        return null;
    }
}
