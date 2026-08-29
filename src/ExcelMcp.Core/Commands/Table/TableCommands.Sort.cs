using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table sort operations.
/// </summary>
public partial class TableCommands
{
    private const int xlYes = 1;

    /// <inheritdoc />
    public OperationResult Sort(
        IExcelBatch batch,
        string tableName,
        string columnName,
        bool ascending = true) =>
        Sort(batch, tableName, columnName, ascending, validateIntegrity: false);

    /// <inheritdoc />
    public TableSortResult Sort(
        IExcelBatch batch,
        string tableName,
        string columnName,
        bool ascending = true,
        bool validateIntegrity = false,
        List<string>? keyColumns = null,
        List<TableSortControlTotal>? controlTotals = null)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(columnName);

        return SortCore(
            batch,
            tableName,
            [new TableSortColumn { ColumnName = columnName, Ascending = ascending }],
            validateIntegrity,
            keyColumns,
            controlTotals);
    }

    /// <inheritdoc />
    public OperationResult SortMulti(
        IExcelBatch batch,
        string tableName,
        List<TableSortColumn> sortColumns) =>
        SortMulti(batch, tableName, sortColumns, validateIntegrity: false);

    /// <inheritdoc />
    public TableSortResult SortMulti(
        IExcelBatch batch,
        string tableName,
        List<TableSortColumn> sortColumns,
        bool validateIntegrity = false,
        List<string>? keyColumns = null,
        List<TableSortControlTotal>? controlTotals = null)
    {
        ArgumentNullException.ThrowIfNull(sortColumns);
        if (sortColumns.Count == 0)
        {
            throw new ArgumentException("At least one sort column must be specified.", nameof(sortColumns));
        }

        if (sortColumns.Count > 3)
        {
            throw new ArgumentException("Excel supports a maximum of 3 sort levels.", nameof(sortColumns));
        }

        foreach (TableSortColumn sortColumn in sortColumns)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(sortColumn.ColumnName);
        }

        return SortCore(
            batch,
            tableName,
            sortColumns,
            validateIntegrity,
            keyColumns,
            controlTotals);
    }

    private static TableSortResult SortCore(
        IExcelBatch batch,
        string tableName,
        IReadOnlyList<TableSortColumn> sortColumns,
        bool validateIntegrity,
        IReadOnlyList<string>? keyColumns,
        IReadOnlyList<TableSortControlTotal>? controlTotals)
    {
        ValidateTableName(tableName);
        ValidateIntegrityArguments(keyColumns, controlTotals);

        bool validationRequested = validateIntegrity
            || keyColumns is { Count: > 0 }
            || controlTotals is { Count: > 0 };
        IReadOnlyList<string> requestedKeyColumns = keyColumns ?? [];
        IReadOnlyList<TableSortControlTotal> requestedControlTotals = controlTotals ?? [];

        return batch.Execute((ctx, ct) =>
        {
            Excel.ListObject? table = null;
            Excel.Range? tableRange = null;
            var keyRanges = new List<Excel.Range>();
            try
            {
                table = (Excel.ListObject)FindTable(ctx.Book, tableName);
                tableRange = table.Range;
                ResolveSortKeyRanges(table, sortColumns, keyRanges);

                var result = new TableSortResult
                {
                    FilePath = batch.WorkbookPath,
                    TableName = tableName,
                    TableRange = Convert.ToString(tableRange.Address) ?? string.Empty,
                    ValidationPerformed = validationRequested
                };

                if (!validationRequested)
                {
                    ApplySort(tableRange, keyRanges, sortColumns);
                    result.Success = true;
                    result.SortAttempted = true;
                    result.SortCommitted = true;
                    return result;
                }

                TableSortSnapshot snapshot = CaptureSortSnapshot(
                    table,
                    tableRange,
                    requestedKeyColumns,
                    requestedControlTotals,
                    result,
                    ct);
                if (result.Findings.Any(finding => finding.Severity == TablePreflightSeverity.Blocker))
                {
                    result.Success = false;
                    result.ErrorMessage = "The table was not sorted because an integrity check could not be evaluated safely.";
                    return result;
                }

                Exception? postSortException = null;
                try
                {
                    ApplySort(tableRange, keyRanges, sortColumns);
                    result.SortAttempted = true;

                    TableSortState postSortState = CaptureTableSortState(
                        table,
                        calculateValues: requestedControlTotals.Count > 0,
                        ct);
                    bool integrityPreserved = ValidatePostSortIntegrity(
                        snapshot,
                        postSortState,
                        requestedControlTotals,
                        result);
                    result.IntegrityPreserved = integrityPreserved;

                    if (integrityPreserved)
                    {
                        result.Success = true;
                        result.SortCommitted = true;
                        return result;
                    }
                }
#pragma warning disable CA1031 // Any failure after mutation must pass through rollback before preserving the original exception.
                catch (Exception ex)
#pragma warning restore CA1031
                {
                    postSortException = ex;
                    result.IntegrityPreserved = false;
                }

                result.RollbackAttempted = true;
                string? rollbackError = null;
                try
                {
                    RestoreSortSnapshot(table, snapshot);
                    TableSortState restoredState = CaptureTableSortState(
                        table,
                        calculateValues: requestedControlTotals.Count > 0,
                        CancellationToken.None);
                    result.RollbackSucceeded = SnapshotWasRestored(snapshot, restoredState);
                }
#pragma warning disable CA1031 // Rollback must report an uncertain workbook state instead of hiding the original failed validation.
                catch (Exception ex) when (ex is not OperationCanceledException)
#pragma warning restore CA1031
                {
                    result.RollbackSucceeded = false;
                    rollbackError = ex.Message;
                }

                result.Success = false;
                result.SortCommitted = false;
                if (postSortException is not null)
                {
                    if (result.RollbackSucceeded == true)
                    {
                        System.Runtime.ExceptionServices.ExceptionDispatchInfo
                            .Capture(postSortException)
                            .Throw();
                    }

                    throw new InvalidOperationException(
                        "Table sorting failed after mutation, and the original table contents could not be restored and verified."
                            + (string.IsNullOrWhiteSpace(rollbackError) ? string.Empty : $" Rollback error: {rollbackError}"),
                        postSortException);
                }

                result.ErrorMessage = result.RollbackSucceeded == true
                    ? "Post-sort integrity validation failed. The original table contents were restored and verified."
                    : "Post-sort integrity validation failed, and the original table contents could not be restored and verified."
                        + (string.IsNullOrWhiteSpace(rollbackError) ? string.Empty : $" Rollback error: {rollbackError}");
                return result;
            }
            finally
            {
                for (int index = keyRanges.Count - 1; index >= 0; index--)
                {
                    Excel.Range? keyRange = keyRanges[index];
                    ComUtilities.Release(ref keyRange);
                }

                ComUtilities.Release(ref tableRange);
                ComUtilities.Release(ref table);
            }
        });
    }

    private static void ValidateIntegrityArguments(
        IReadOnlyList<string>? keyColumns,
        IReadOnlyList<TableSortControlTotal>? controlTotals)
    {
        if (keyColumns is not null)
        {
            foreach (string keyColumn in keyColumns)
            {
                ArgumentException.ThrowIfNullOrWhiteSpace(keyColumn);
            }

            if (keyColumns.Distinct(StringComparer.OrdinalIgnoreCase).Count() != keyColumns.Count)
            {
                throw new ArgumentException("Row-key column names must be unique.", nameof(keyColumns));
            }
        }

        if (controlTotals is null)
        {
            return;
        }

        foreach (TableSortControlTotal controlTotal in controlTotals)
        {
            ArgumentNullException.ThrowIfNull(controlTotal);
            ArgumentException.ThrowIfNullOrWhiteSpace(controlTotal.ColumnName);
            if (controlTotal.Tolerance < 0)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(controlTotals),
                    "Control-total tolerance cannot be negative.");
            }
        }

        if (controlTotals
            .Select(total => total.ColumnName)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count() != controlTotals.Count)
        {
            throw new ArgumentException("Control-total column names must be unique.", nameof(controlTotals));
        }
    }

    private static void ResolveSortKeyRanges(
        Excel.ListObject table,
        IReadOnlyList<TableSortColumn> sortColumns,
        List<Excel.Range> keyRanges)
    {
        Excel.ListColumns? columns = null;
        try
        {
            columns = table.ListColumns;
            foreach (TableSortColumn sortColumn in sortColumns)
            {
                Excel.ListColumn? matchedColumn = null;
                try
                {
                    for (int index = 1; index <= columns.Count; index++)
                    {
                        Excel.ListColumn? candidate = null;
                        try
                        {
                            candidate = columns.Item[index];
                            if (string.Equals(
                                candidate.Name,
                                sortColumn.ColumnName,
                                StringComparison.OrdinalIgnoreCase))
                            {
                                matchedColumn = candidate;
                                candidate = null;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref candidate);
                        }
                    }

                    if (matchedColumn is null)
                    {
                        throw new InvalidOperationException(
                            $"Column '{sortColumn.ColumnName}' not found in table '{table.Name}'.");
                    }

                    keyRanges.Add(matchedColumn.Range);
                }
                finally
                {
                    ComUtilities.Release(ref matchedColumn);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref columns);
        }
    }

    private static void ApplySort(
        Excel.Range sortRange,
        List<Excel.Range> keyRanges,
        IReadOnlyList<TableSortColumn> sortColumns)
    {
        // Reason: Excel's early-bound _Sort overload rejects valid ListObject range keys that the late-bound COM call accepts.
        dynamic sortTarget = sortRange;
        if (sortColumns.Count == 1)
        {
            sortTarget.Sort(
                Key1: keyRanges[0],
                Order1: sortColumns[0].Ascending ? 1 : 2,
                Header: xlYes);
            return;
        }

        if (sortColumns.Count == 2)
        {
            sortTarget.Sort(
                Key1: keyRanges[0],
                Order1: sortColumns[0].Ascending ? 1 : 2,
                Key2: keyRanges[1],
                Order2: sortColumns[1].Ascending ? 1 : 2,
                Header: xlYes);
            return;
        }

        sortTarget.Sort(
            Key1: keyRanges[0],
            Order1: sortColumns[0].Ascending ? 1 : 2,
            Key2: keyRanges[1],
            Order2: sortColumns[1].Ascending ? 1 : 2,
            Key3: keyRanges[2],
            Order3: sortColumns[2].Ascending ? 1 : 2,
            Header: xlYes);
    }
}
