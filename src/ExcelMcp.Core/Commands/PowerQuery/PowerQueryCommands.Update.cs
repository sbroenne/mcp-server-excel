using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Formatting;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Power Query Update operations.
/// Reuses the shared refresh helper so update and refresh follow the same COM-safe path.
/// </summary>
public partial class PowerQueryCommands
{
    /// <summary>
    /// Update Power Query M code. Preserves load configuration (worksheet/data model).
    /// M code is automatically formatted using the powerqueryformatter.com API before saving.
    /// - Worksheet queries: Uses QueryTable.Refresh(false) for synchronous refresh with column propagation
    /// - Data Model queries: Uses connection.Refresh() to update the Data Model
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="queryName">Name of query to update</param>
    /// <param name="mCode">New M code</param>
    /// <param name="refresh">Whether to refresh data after update (default: true)</param>
    /// <exception cref="ArgumentException">Thrown when queryName or mCode is invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when query not found or update fails</exception>
    public OperationResult Update(IExcelBatch batch, string queryName, string mCode, bool refresh = true)
    {
        if (!ValidateQueryName(queryName, out string? validationError))
        {
            throw new ArgumentException(validationError, nameof(queryName));
        }

        if (string.IsNullOrWhiteSpace(mCode))
        {
            throw new ArgumentException("M code cannot be empty", nameof(mCode));
        }

        // Format M code before saving (outside batch.Execute for async operation)
        // Formatting is done synchronously to maintain method signature compatibility
        // Falls back to original if formatting fails
        string formattedMCode = MCodeFormatter.FormatAsync(mCode).GetAwaiter().GetResult();

        return batch.Execute((ctx, ct) =>
        {
            Excel.Queries? queries = null;
            Excel.WorkbookQuery? query = null;

            try
            {
                // STEP 1: Find the Power Query
                queries = ctx.Book.Queries;
                query = null;
                for (int i = 1; i <= queries.Count; i++)
                {
                    dynamic? q = null;
                    try
                    {
                        q = queries.Item(i);
                        string qName = q.Name?.ToString() ?? "";
                        if (qName.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                        {
                            query = q;
                            q = null; // Don't release - we're keeping the reference
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref q!);
                    }
                }

                if (query == null)
                {
                    throw new InvalidOperationException($"Query '{queryName}' not found.");
                }

                // STEP 2: Update the M code with formatted version
                // Note: 0x800A03EC error can occur in certain workbook states (see Issue #323)
                // Retry doesn't help - it's a workbook state issue, not transient
                query.Formula = formattedMCode;

                // STEP 3: Refresh if requested using the same COM-safe helper as Refresh().
                // This keeps Update aligned with the message-filter and cancellation behavior
                // already hardened for synchronous worksheet and data model refresh paths.
                if (refresh)
                {
                    _ = RefreshConnectionByQueryName(ctx.Book, queryName, ct);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref query!);
                ComUtilities.Release(ref queries!);
            }
        });
    }

}


