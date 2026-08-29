using System.Globalization;
using System.Text;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

public partial class TableCommands
{
    private static TableSortSnapshot CaptureSortSnapshot(
        Excel.ListObject table,
        Excel.Range tableRange,
        IReadOnlyList<string> keyColumns,
        IReadOnlyList<TableSortControlTotal> controlTotals,
        TableSortResult result,
        CancellationToken cancellationToken)
    {
        TableSortState state = CaptureTableSortState(
            table,
            calculateValues: controlTotals.Count > 0,
            cancellationToken);
        var snapshot = new TableSortSnapshot { State = state };

        Excel.Range? dataBodyRange = null;
        try
        {
            dataBodyRange = table.DataBodyRange;
            if (dataBodyRange is not null)
            {
                InspectSortSensitiveFormulas(
                    dataBodyRange,
                    hasHeaders: false,
                    result.Findings,
                    cancellationToken);
            }

            InspectAdjacentTableDataAndFormulas(
                tableRange,
                dataBodyRange,
                result.Findings,
                cancellationToken);
        }
        finally
        {
            ComUtilities.Release(ref dataBodyRange);
        }

        snapshot.CalculatedColumns = AnalyzeCalculatedColumns(state, result.Findings);

        if (!TryResolveColumnIndexes(state.Headers, keyColumns, out int[] keyIndexes, out string? missingKey))
        {
            result.Findings.Add(new TablePreflightFinding
            {
                Kind = TablePreflightFindingKind.InvalidRowKey,
                Severity = TablePreflightSeverity.Blocker,
                Message = $"Row-key column '{missingKey}' was not found in the table.",
                Remediation = "Choose existing table columns for the composite row key."
            });
        }
        else if (keyIndexes.Length > 0)
        {
            if (!TryBuildKeyedRows(
                state,
                keyIndexes,
                out Dictionary<string, KeyedRow> keyedRows,
                out TablePreflightFindingKind failureKind,
                out string failureMessage))
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = failureKind,
                    Severity = TablePreflightSeverity.Blocker,
                    Message = failureMessage,
                    Remediation = "Choose columns whose combined values are populated and unique for every table row."
                });
            }
            else
            {
                snapshot.KeyColumnIndexes = keyIndexes;
                snapshot.KeyedRows = keyedRows;
            }
        }

        foreach (TableSortControlTotal controlTotal in controlTotals)
        {
            int columnIndex = FindColumnIndex(state.Headers, controlTotal.ColumnName);
            if (columnIndex < 0)
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.InvalidControlTotal,
                    Severity = TablePreflightSeverity.Blocker,
                    Message = $"Control-total column '{controlTotal.ColumnName}' was not found in the table.",
                    Remediation = "Choose an existing numeric table column for the control total."
                });
                continue;
            }

            if (!TryCalculateControlTotal(state, columnIndex, out decimal total, out string? error))
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.InvalidControlTotal,
                    Severity = TablePreflightSeverity.Blocker,
                    Message = $"Control-total column '{controlTotal.ColumnName}' cannot be summed: {error}",
                    Remediation = "Remove nonnumeric or error values, or choose another numeric column."
                });
                continue;
            }

            snapshot.ControlTotals.Add(new ControlTotalSnapshot
            {
                ColumnIndex = columnIndex,
                Request = controlTotal,
                Before = total
            });
        }

        return snapshot;
    }

    private static TableSortState CaptureTableSortState(
        Excel.ListObject table,
        bool calculateValues,
        CancellationToken cancellationToken)
    {
        Excel.Range? tableRange = null;
        Excel.Range? tableRows = null;
        Excel.Range? tableColumns = null;
        Excel.Range? dataBodyRange = null;
        Excel.Range? dataRows = null;
        Excel.Range? totalsRowRange = null;
        Excel.ListColumns? listColumns = null;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            tableRange = table.Range;
            tableRows = tableRange.Rows;
            tableColumns = tableRange.Columns;
            dataBodyRange = table.DataBodyRange;
            if (calculateValues && dataBodyRange is not null)
            {
                dataBodyRange.Calculate();
            }

            int tableRowCount = Convert.ToInt32(tableRows.Count, CultureInfo.InvariantCulture);
            int columnCount = Convert.ToInt32(tableColumns.Count, CultureInfo.InvariantCulture);
            int dataRowCount = 0;
            int dataFirstRow = 0;
            object? dataFormulas = null;
            object? dataValues = null;
            bool[,] dataFormulaFlags = new bool[0, columnCount];
            if (dataBodyRange is not null)
            {
                dataRows = dataBodyRange.Rows;
                dataRowCount = Convert.ToInt32(dataRows.Count, CultureInfo.InvariantCulture);
                dataFirstRow = Convert.ToInt32(dataBodyRange.Row, CultureInfo.InvariantCulture);
                dataFormulas = CloneMatrix(dataBodyRange.FormulaR1C1);
                dataValues = CloneMatrix(dataBodyRange.Value2);
                dataFormulaFlags = CaptureFormulaFlags(dataBodyRange, dataRowCount, columnCount);
            }

            bool showTotals = table.ShowTotals;
            object? totalsContent = null;
            object? totalsValues = null;
            bool[,] totalsFormulaFlags = new bool[0, columnCount];
            if (showTotals)
            {
                totalsRowRange = table.TotalsRowRange;
                totalsContent = CloneMatrix(totalsRowRange.FormulaR1C1);
                totalsValues = CloneMatrix(totalsRowRange.Value2);
                totalsFormulaFlags = CaptureFormulaFlags(totalsRowRange, 1, columnCount);
            }

            listColumns = table.ListColumns;
            var headers = new List<string>(columnCount);
            for (int index = 1; index <= listColumns.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.ListColumn? column = null;
                try
                {
                    column = listColumns.Item[index];
                    headers.Add(column.Name);
                }
                finally
                {
                    ComUtilities.Release(ref column);
                }
            }

            return new TableSortState
            {
                RangeAddress = Convert.ToString(tableRange.Address) ?? string.Empty,
                TableRowCount = tableRowCount,
                ColumnCount = columnCount,
                TableFirstColumn = Convert.ToInt32(tableRange.Column, CultureInfo.InvariantCulture),
                DataFirstRow = dataFirstRow,
                DataRowCount = dataRowCount,
                Headers = headers,
                ShowTotals = showTotals,
                TotalsContentSignature = BuildContentSignature(
                    totalsContent,
                    totalsValues,
                    totalsFormulaFlags,
                    showTotals ? 1 : 0,
                    columnCount),
                TotalsFormulaMatrix = totalsContent,
                TotalsValueMatrix = totalsValues,
                TotalsFormulaFlags = totalsFormulaFlags,
                DataFormulaMatrix = dataFormulas,
                DataValueMatrix = dataValues,
                DataFormulaFlags = dataFormulaFlags,
                DataContentSignature = BuildContentSignature(
                    dataFormulas,
                    dataValues,
                    dataFormulaFlags,
                    dataRowCount,
                    columnCount),
                RowSignatures = BuildRowSignatures(
                    dataFormulas,
                    dataValues,
                    dataFormulaFlags,
                    dataRowCount,
                    columnCount)
            };
        }
        finally
        {
            ComUtilities.Release(ref listColumns);
            ComUtilities.Release(ref totalsRowRange);
            ComUtilities.Release(ref dataRows);
            ComUtilities.Release(ref dataBodyRange);
            ComUtilities.Release(ref tableColumns);
            ComUtilities.Release(ref tableRows);
            ComUtilities.Release(ref tableRange);
        }
    }

    private static List<CalculatedColumnSnapshot> AnalyzeCalculatedColumns(
        TableSortState state,
        List<TablePreflightFinding> findings)
    {
        var calculatedColumns = new List<CalculatedColumnSnapshot>();
        if (state.DataRowCount == 0)
        {
            return calculatedColumns;
        }

        for (int columnIndex = 0; columnIndex < state.ColumnCount; columnIndex++)
        {
            var formulas = new List<string>();
            int formulaCount = 0;
            for (int rowIndex = 0; rowIndex < state.DataRowCount; rowIndex++)
            {
                if (state.DataFormulaFlags[rowIndex, columnIndex])
                {
                    formulaCount++;
                    formulas.Add(Convert.ToString(
                        GetMatrixValue(state.DataFormulaMatrix ?? string.Empty, rowIndex, columnIndex),
                        CultureInfo.InvariantCulture) ?? string.Empty);
                }
            }

            if (formulaCount == 0)
            {
                continue;
            }

            string[] patterns = formulas.Distinct(StringComparer.Ordinal).ToArray();
            if (formulaCount == state.DataRowCount && patterns.Length == 1)
            {
                calculatedColumns.Add(new CalculatedColumnSnapshot
                {
                    ColumnIndex = columnIndex,
                    ColumnName = state.Headers[columnIndex],
                    FormulaR1C1 = patterns[0]
                });
                continue;
            }

            findings.Add(new TablePreflightFinding
            {
                Kind = TablePreflightFindingKind.MixedFormulaColumn,
                Severity = TablePreflightSeverity.Warning,
                IsHeuristic = true,
                Addresses =
                [
                    GetAbsoluteRangeAddress(
                        state.TableFirstColumn + columnIndex,
                        state.DataFirstRow,
                        state.TableFirstColumn + columnIndex,
                        state.DataFirstRow + state.DataRowCount - 1)
                ],
                Message = $"Column '{state.Headers[columnIndex]}' mixes formulas, formula patterns, or literal values and may contain intentional overrides.",
                Remediation = "Confirm that the mixed cells are intentional before relying on calculated-column consistency."
            });
        }

        return calculatedColumns;
    }

    private static bool ValidatePostSortIntegrity(
        TableSortSnapshot snapshot,
        TableSortState postSortState,
        IReadOnlyList<TableSortControlTotal> requestedControlTotals,
        TableSortResult result)
    {
        TableSortState before = snapshot.State;
        TableSortIntegrityChecks checks = result.Checks;
        checks.RangePreserved = string.Equals(
            before.RangeAddress,
            postSortState.RangeAddress,
            StringComparison.Ordinal);
        checks.ShapePreserved = before.TableRowCount == postSortState.TableRowCount
            && before.DataRowCount == postSortState.DataRowCount
            && before.ColumnCount == postSortState.ColumnCount;
        checks.HeadersPreserved = before.Headers.SequenceEqual(
            postSortState.Headers,
            StringComparer.Ordinal);
        checks.TotalsRowPreserved = before.ShowTotals == postSortState.ShowTotals
            && string.Equals(
                before.TotalsContentSignature,
                postSortState.TotalsContentSignature,
                StringComparison.Ordinal);
        checks.RowSetPreserved = MultisetEquals(before.RowSignatures, postSortState.RowSignatures);

        if (checks.RangePreserved != true
            || checks.ShapePreserved != true
            || checks.HeadersPreserved != true
            || checks.TotalsRowPreserved != true)
        {
            result.Findings.Add(new TablePreflightFinding
            {
                Kind = TablePreflightFindingKind.TableStructureChanged,
                Severity = TablePreflightSeverity.Blocker,
                Addresses = [before.RangeAddress],
                Message = "The table range, shape, headers, or totals row changed during sorting.",
                Remediation = "Review the table structure after the operation; automatic rollback has been attempted."
            });
        }

        if (checks.RowSetPreserved != true)
        {
            result.Findings.Add(new TablePreflightFinding
            {
                Kind = TablePreflightFindingKind.TableRowsChanged,
                Severity = TablePreflightSeverity.Blocker,
                Addresses = [before.RangeAddress],
                Message = "Complete logical table rows were not preserved as a permutation.",
                Remediation = "Review formulas and row data; automatic rollback has been attempted."
            });
        }

        foreach (CalculatedColumnSnapshot calculatedColumn in snapshot.CalculatedColumns)
        {
            bool consistentAfter = ColumnMatchesFormula(postSortState, calculatedColumn);
            checks.CalculatedColumns.Add(new TableCalculatedColumnCheckResult
            {
                ColumnName = calculatedColumn.ColumnName,
                FormulaR1C1 = calculatedColumn.FormulaR1C1,
                ConsistentBefore = true,
                ConsistentAfter = consistentAfter,
                Passed = consistentAfter
            });
            if (!consistentAfter)
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.CalculatedColumnChanged,
                    Severity = TablePreflightSeverity.Blocker,
                    Message = $"Calculated column '{calculatedColumn.ColumnName}' no longer has its original formula pattern.",
                    Remediation = "Review the calculated-column formula; automatic rollback has been attempted."
                });
            }
        }

        if (snapshot.KeyColumnIndexes.Length > 0)
        {
            var keyCheck = new TableRowKeyCheckResult
            {
                KeyColumns = snapshot.KeyColumnIndexes.Select(index => before.Headers[index]).ToList(),
                BeforeCount = snapshot.KeyedRows.Count
            };
            if (!TryBuildKeyedRows(
                postSortState,
                snapshot.KeyColumnIndexes,
                out Dictionary<string, KeyedRow> afterRows,
                out _,
                out _))
            {
                keyCheck.Passed = false;
            }
            else
            {
                keyCheck.AfterCount = afterRows.Count;
                foreach ((string key, KeyedRow originalRow) in snapshot.KeyedRows)
                {
                    if (!afterRows.TryGetValue(key, out KeyedRow? afterRow)
                        || !string.Equals(originalRow.RowSignature, afterRow.RowSignature, StringComparison.Ordinal))
                    {
                        keyCheck.MismatchedKeys.Add(originalRow.DisplayKey);
                    }
                }

                foreach ((string key, KeyedRow afterRow) in afterRows)
                {
                    if (!snapshot.KeyedRows.ContainsKey(key))
                    {
                        keyCheck.MismatchedKeys.Add(afterRow.DisplayKey);
                    }
                }

                keyCheck.Passed = keyCheck.BeforeCount == keyCheck.AfterCount
                    && keyCheck.MismatchedKeys.Count == 0;
            }

            checks.RowKeys = keyCheck;
            if (!keyCheck.Passed)
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.RowKeyMismatch,
                    Severity = TablePreflightSeverity.Blocker,
                    Message = "Caller-supplied row identities or their complete row contents changed during sorting.",
                    Remediation = "Review the reported key columns; automatic rollback has been attempted."
                });
            }
        }

        foreach (ControlTotalSnapshot controlTotal in snapshot.ControlTotals)
        {
            bool calculated = TryCalculateControlTotal(
                postSortState,
                controlTotal.ColumnIndex,
                out decimal after,
                out _);
            decimal delta = calculated ? after - controlTotal.Before : decimal.MaxValue;
            bool passed = calculated && Math.Abs(delta) <= controlTotal.Request.Tolerance;
            checks.ControlTotals.Add(new TableControlTotalCheckResult
            {
                ColumnName = controlTotal.Request.ColumnName,
                Before = controlTotal.Before,
                After = calculated ? after : controlTotal.Before,
                Delta = calculated ? delta : 0,
                Tolerance = controlTotal.Request.Tolerance,
                Passed = passed
            });
            if (!passed)
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.ControlTotalMismatch,
                    Severity = TablePreflightSeverity.Blocker,
                    Message = $"Control total for column '{controlTotal.Request.ColumnName}' changed beyond the requested tolerance.",
                    Remediation = "Review row-dependent formulas or values; automatic rollback has been attempted."
                });
            }
        }

        return checks.RangePreserved == true
            && checks.ShapePreserved == true
            && checks.HeadersPreserved == true
            && checks.TotalsRowPreserved == true
            && checks.RowSetPreserved == true
            && checks.CalculatedColumns.All(check => check.Passed)
            && (checks.RowKeys?.Passed ?? true)
            && checks.ControlTotals.All(check => check.Passed)
            && requestedControlTotals.Count == checks.ControlTotals.Count;
    }

    private static void RestoreSortSnapshot(Excel.ListObject table, TableSortSnapshot snapshot)
    {
        Excel.ListColumns? listColumns = null;
        Excel.Range? dataBodyRange = null;
        Excel.Range? totalsRowRange = null;
        try
        {
            listColumns = table.ListColumns;
            if (listColumns.Count != snapshot.State.ColumnCount)
            {
                throw new InvalidOperationException("The table column count changed, so the captured contents cannot be restored.");
            }

            for (int index = 1; index <= listColumns.Count; index++)
            {
                Excel.ListColumn? column = null;
                try
                {
                    column = listColumns.Item[index];
                    column.Name = snapshot.State.Headers[index - 1];
                }
                finally
                {
                    ComUtilities.Release(ref column);
                }
            }

            dataBodyRange = table.DataBodyRange;
            if (snapshot.State.DataRowCount > 0)
            {
                if (dataBodyRange is null)
                {
                    throw new InvalidOperationException("The table data body is missing, so the captured contents cannot be restored.");
                }

                RestoreRangeContent(
                    dataBodyRange,
                    snapshot.State.DataFormulaMatrix,
                    snapshot.State.DataValueMatrix,
                    snapshot.State.DataFormulaFlags,
                    snapshot.State.DataRowCount,
                    snapshot.State.ColumnCount);
            }

            if (snapshot.State.ShowTotals)
            {
                if (!table.ShowTotals)
                {
                    table.ShowTotals = true;
                }

                totalsRowRange = table.TotalsRowRange;
                RestoreRangeContent(
                    totalsRowRange,
                    snapshot.State.TotalsFormulaMatrix,
                    snapshot.State.TotalsValueMatrix,
                    snapshot.State.TotalsFormulaFlags,
                    1,
                    snapshot.State.ColumnCount);
            }
            else if (table.ShowTotals)
            {
                table.ShowTotals = false;
            }
        }
        finally
        {
            ComUtilities.Release(ref totalsRowRange);
            ComUtilities.Release(ref dataBodyRange);
            ComUtilities.Release(ref listColumns);
        }
    }

    private static void RestoreRangeContent(
        Excel.Range range,
        object? formulaMatrix,
        object? valueMatrix,
        bool[,] formulaFlags,
        int rowCount,
        int columnCount)
    {
        range.Value2 = valueMatrix;
        Excel.Range? cells = null;
        try
        {
            cells = range.Cells;
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    object? value = GetMatrixValue(
                        valueMatrix ?? string.Empty,
                        rowIndex,
                        columnIndex);
                    bool restoreFormula = formulaFlags[rowIndex, columnIndex];
                    bool forceLiteralText = !restoreFormula
                        && value is string text
                        && text.Length > 0
                        && text[0] is '=' or '+' or '-' or '@';
                    if (!restoreFormula && !forceLiteralText)
                    {
                        continue;
                    }

                    Excel.Range? cell = null;
                    try
                    {
                        cell = cells[rowIndex + 1, columnIndex + 1];
                        cell.FormulaR1C1 = restoreFormula
                            ? Convert.ToString(
                                GetMatrixValue(formulaMatrix ?? string.Empty, rowIndex, columnIndex),
                                CultureInfo.InvariantCulture) ?? string.Empty
                            : $"'{value}";
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                    }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref cells);
        }
    }

    private static bool SnapshotWasRestored(
        TableSortSnapshot snapshot,
        TableSortState restoredState) =>
        string.Equals(snapshot.State.RangeAddress, restoredState.RangeAddress, StringComparison.Ordinal)
        && snapshot.State.TableRowCount == restoredState.TableRowCount
        && snapshot.State.ColumnCount == restoredState.ColumnCount
        && snapshot.State.Headers.SequenceEqual(restoredState.Headers, StringComparer.Ordinal)
        && snapshot.State.ShowTotals == restoredState.ShowTotals
        && string.Equals(
            snapshot.State.DataContentSignature,
            restoredState.DataContentSignature,
            StringComparison.Ordinal)
        && string.Equals(
            snapshot.State.TotalsContentSignature,
            restoredState.TotalsContentSignature,
            StringComparison.Ordinal);

    private static bool ColumnMatchesFormula(
        TableSortState state,
        CalculatedColumnSnapshot calculatedColumn)
    {
        if (calculatedColumn.ColumnIndex >= state.ColumnCount)
        {
            return false;
        }

        for (int rowIndex = 0; rowIndex < state.DataRowCount; rowIndex++)
        {
            if (!state.DataFormulaFlags[rowIndex, calculatedColumn.ColumnIndex]
                || !string.Equals(
                    Convert.ToString(
                        GetMatrixValue(
                            state.DataFormulaMatrix ?? string.Empty,
                            rowIndex,
                            calculatedColumn.ColumnIndex),
                        CultureInfo.InvariantCulture),
                    calculatedColumn.FormulaR1C1,
                    StringComparison.Ordinal))
            {
                return false;
            }
        }

        return true;
    }

    private static bool TryResolveColumnIndexes(
        IReadOnlyList<string> headers,
        IReadOnlyList<string> requestedColumns,
        out int[] indexes,
        out string? missingColumn)
    {
        indexes = new int[requestedColumns.Count];
        for (int index = 0; index < requestedColumns.Count; index++)
        {
            int columnIndex = FindColumnIndex(headers, requestedColumns[index]);
            if (columnIndex < 0)
            {
                indexes = [];
                missingColumn = requestedColumns[index];
                return false;
            }

            indexes[index] = columnIndex;
        }

        missingColumn = null;
        return true;
    }

    private static int FindColumnIndex(IReadOnlyList<string> headers, string columnName)
    {
        for (int index = 0; index < headers.Count; index++)
        {
            if (string.Equals(headers[index], columnName, StringComparison.OrdinalIgnoreCase))
            {
                return index;
            }
        }

        return -1;
    }

    private static bool TryBuildKeyedRows(
        TableSortState state,
        IReadOnlyList<int> keyColumnIndexes,
        out Dictionary<string, KeyedRow> keyedRows,
        out TablePreflightFindingKind failureKind,
        out string failureMessage)
    {
        keyedRows = new Dictionary<string, KeyedRow>(StringComparer.Ordinal);
        for (int rowIndex = 0; rowIndex < state.DataRowCount; rowIndex++)
        {
            var canonicalKey = new StringBuilder();
            var displayKey = new StringBuilder();
            foreach (int columnIndex in keyColumnIndexes)
            {
                object? value = GetMatrixValue(
                    state.DataValueMatrix ?? string.Empty,
                    rowIndex,
                    columnIndex);
                if (value is null || value is string text && string.IsNullOrWhiteSpace(text))
                {
                    failureKind = TablePreflightFindingKind.InvalidRowKey;
                    failureMessage = "One or more composite row keys are blank.";
                    return false;
                }

                AppendLengthPrefixed(canonicalKey, CanonicalizeCell(value));
                if (displayKey.Length > 0)
                {
                    displayKey.Append(" | ");
                }

                displayKey.Append(Convert.ToString(value, CultureInfo.InvariantCulture));
            }

            string key = canonicalKey.ToString();
            if (!keyedRows.TryAdd(
                key,
                new KeyedRow
                {
                    DisplayKey = displayKey.ToString(),
                    RowSignature = state.RowSignatures[rowIndex]
                }))
            {
                failureKind = TablePreflightFindingKind.DuplicateRowKey;
                failureMessage = $"Composite row key '{displayKey}' is not unique.";
                return false;
            }
        }

        failureKind = default;
        failureMessage = string.Empty;
        return true;
    }

    private static bool TryCalculateControlTotal(
        TableSortState state,
        int columnIndex,
        out decimal total,
        out string? error)
    {
        total = 0;
        try
        {
            for (int rowIndex = 0; rowIndex < state.DataRowCount; rowIndex++)
            {
                object? value = GetMatrixValue(
                    state.DataValueMatrix ?? string.Empty,
                    rowIndex,
                    columnIndex);
                if (value is null || value is string text && string.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                if (!TryConvertNumericValue(value, out decimal number))
                {
                    error = $"row {rowIndex + 1} contains a nonnumeric or error value";
                    return false;
                }

                total = checked(total + number);
            }
        }
        catch (OverflowException)
        {
            error = "the numeric sum exceeds the supported decimal range";
            return false;
        }

        error = null;
        return true;
    }

    private static bool TryConvertNumericValue(object value, out decimal number)
    {
        switch (value)
        {
            case byte or sbyte or short or ushort or int or uint or long or ulong or decimal:
                number = Convert.ToDecimal(value, CultureInfo.InvariantCulture);
                return true;
            case float single when !float.IsNaN(single) && !float.IsInfinity(single):
                number = Convert.ToDecimal(single, CultureInfo.InvariantCulture);
                return true;
            case double @double when !double.IsNaN(@double) && !double.IsInfinity(@double):
                number = Convert.ToDecimal(@double, CultureInfo.InvariantCulture);
                return true;
            default:
                number = 0;
                return false;
        }
    }

    private static List<string> BuildRowSignatures(
        object? dataFormulaMatrix,
        object? dataValueMatrix,
        bool[,] formulaFlags,
        int rowCount,
        int columnCount)
    {
        var signatures = new List<string>(rowCount);
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            var signature = new StringBuilder();
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                AppendLengthPrefixed(
                    signature,
                    CanonicalizeLogicalCell(
                        GetMatrixValue(dataFormulaMatrix ?? string.Empty, rowIndex, columnIndex),
                        GetMatrixValue(dataValueMatrix ?? string.Empty, rowIndex, columnIndex),
                        formulaFlags[rowIndex, columnIndex]));
            }

            signatures.Add(signature.ToString());
        }

        return signatures;
    }

    private static bool MultisetEquals(List<string> before, List<string> after)
    {
        if (before.Count != after.Count)
        {
            return false;
        }

        var counts = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (string signature in before)
        {
            counts[signature] = counts.GetValueOrDefault(signature) + 1;
        }

        foreach (string signature in after)
        {
            if (!counts.TryGetValue(signature, out int count))
            {
                return false;
            }

            if (count == 1)
            {
                counts.Remove(signature);
            }
            else
            {
                counts[signature] = count - 1;
            }
        }

        return counts.Count == 0;
    }

    private static string BuildContentSignature(
        object? formulaMatrix,
        object? valueMatrix,
        bool[,] formulaFlags,
        int rowCount,
        int columnCount)
    {
        if (rowCount == 0 || columnCount == 0)
        {
            return string.Empty;
        }

        var signature = new StringBuilder();
        for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
        {
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                AppendLengthPrefixed(
                    signature,
                    CanonicalizeLogicalCell(
                        GetMatrixValue(formulaMatrix ?? string.Empty, rowIndex, columnIndex),
                        GetMatrixValue(valueMatrix ?? string.Empty, rowIndex, columnIndex),
                        formulaFlags[rowIndex, columnIndex]));
            }
        }

        return signature.ToString();
    }

    private static string CanonicalizeCell(object? value) =>
        value switch
        {
            null => "null",
            string text => $"string:{text}",
            bool boolean => boolean ? "bool:true" : "bool:false",
            double number => $"double:{number.ToString("R", CultureInfo.InvariantCulture)}",
            float number => $"float:{number.ToString("R", CultureInfo.InvariantCulture)}",
            decimal number => $"decimal:{number.ToString(CultureInfo.InvariantCulture)}",
            DateTime dateTime => $"datetime:{dateTime.ToString("O", CultureInfo.InvariantCulture)}",
            _ => $"{value.GetType().FullName}:{Convert.ToString(value, CultureInfo.InvariantCulture)}"
        };

    private static string CanonicalizeLogicalCell(
        object? formula,
        object? value,
        bool hasFormula) =>
        hasFormula
            ? $"formula:{Convert.ToString(formula, CultureInfo.InvariantCulture)}"
            : $"literal:{CanonicalizeCell(value)}";

    private static bool[,] CaptureFormulaFlags(
        Excel.Range range,
        int rowCount,
        int columnCount)
    {
        var flags = new bool[rowCount, columnCount];
        Excel.Range? cells = null;
        try
        {
            cells = range.Cells;
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
                {
                    Excel.Range? cell = null;
                    try
                    {
                        cell = cells[rowIndex + 1, columnIndex + 1];
                        flags[rowIndex, columnIndex] = Convert.ToBoolean(
                            cell.HasFormula,
                            CultureInfo.InvariantCulture);
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                    }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref cells);
        }

        return flags;
    }

    private static void AppendLengthPrefixed(StringBuilder builder, string value) =>
        builder.Append(value.Length).Append(':').Append(value);

    private static object? CloneMatrix(object? matrix) =>
        matrix is Array array ? array.Clone() : matrix;

    private sealed class TableSortSnapshot
    {
        public TableSortState State { get; set; } = new();

        public List<CalculatedColumnSnapshot> CalculatedColumns { get; set; } = [];

        public int[] KeyColumnIndexes { get; set; } = [];

        public Dictionary<string, KeyedRow> KeyedRows { get; set; } = new(StringComparer.Ordinal);

        public List<ControlTotalSnapshot> ControlTotals { get; } = [];
    }

    private sealed class TableSortState
    {
        public string RangeAddress { get; set; } = string.Empty;

        public int TableRowCount { get; set; }

        public int ColumnCount { get; set; }

        public int TableFirstColumn { get; set; }

        public int DataFirstRow { get; set; }

        public int DataRowCount { get; set; }

        public List<string> Headers { get; set; } = [];

        public bool ShowTotals { get; set; }

        public string TotalsContentSignature { get; set; } = string.Empty;

        public object? TotalsFormulaMatrix { get; set; }

        public object? TotalsValueMatrix { get; set; }

        public bool[,] TotalsFormulaFlags { get; set; } = new bool[0, 0];

        public object? DataFormulaMatrix { get; set; }

        public object? DataValueMatrix { get; set; }

        public bool[,] DataFormulaFlags { get; set; } = new bool[0, 0];

        public string DataContentSignature { get; set; } = string.Empty;

        public List<string> RowSignatures { get; set; } = [];
    }

    private sealed class CalculatedColumnSnapshot
    {
        public int ColumnIndex { get; set; }

        public string ColumnName { get; set; } = string.Empty;

        public string FormulaR1C1 { get; set; } = string.Empty;
    }

    private sealed class ControlTotalSnapshot
    {
        public int ColumnIndex { get; set; }

        public TableSortControlTotal Request { get; set; } = new();

        public decimal Before { get; set; }
    }

    private sealed class KeyedRow
    {
        public string DisplayKey { get; set; } = string.Empty;

        public string RowSignature { get; set; } = string.Empty;
    }
}
