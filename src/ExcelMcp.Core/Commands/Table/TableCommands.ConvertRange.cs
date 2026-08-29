using System.Globalization;
using System.Runtime.InteropServices;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Atomic range-to-table conversion.
/// </summary>
public partial class TableCommands
{
    private const string RollbackSheetPrefix = "__ExcelMcpTblRb_";

    /// <inheritdoc />
    public TableRangeConversionResult ConvertRange(
        IExcelBatch batch,
        string sheetName,
        string tableName,
        string rangeAddress,
        bool hasHeaders = true,
        string? tableStyle = null,
        TableMergedHeaderPolicy mergedHeaderPolicy = TableMergedHeaderPolicy.Report,
        TableHeaderPolicy headerPolicy = TableHeaderPolicy.Report)
    {
        ValidateCreateInputs(sheetName, tableName, rangeAddress);
        ValidateConversionPolicies(hasHeaders, mergedHeaderPolicy, headerPolicy);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? effectiveRange = null;
            dynamic? listObjects = null;
            dynamic? newTable = null;
            RangeRollbackSnapshot? snapshot = null;
            var stage = TableConversionFailureStage.Preflight;
            var sourceMutationStarted = false;
            var result = new TableRangeConversionResult
            {
                FilePath = batch.WorkbookPath,
                SheetName = sheetName,
                TableName = tableName,
                RequestedRange = rangeAddress,
                MergedHeaderPolicy = mergedHeaderPolicy,
                HeaderPolicy = headerPolicy
            };

            try
            {
                ct.ThrowIfCancellationRequested();
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                effectiveRange = ResolveEffectiveRange(sheet, rangeAddress);
                result.EffectiveRange = effectiveRange.Address;

                TablePreflightResult preflight = AnalyzePreflight(
                    ctx.Book,
                    effectiveRange,
                    batch.WorkbookPath,
                    sheetName,
                    tableName,
                    rangeAddress,
                    hasHeaders,
                    ct);
                AddConversionOnlyBlockers(effectiveRange, preflight, ct);
                result.PreflightFindings = preflight.Findings;

                List<TablePreflightFinding> unhandledBlockers = GetUnhandledConversionBlockers(
                    sheet,
                    effectiveRange,
                    preflight,
                    hasHeaders,
                    mergedHeaderPolicy,
                    headerPolicy,
                    ct);
                if (unhandledBlockers.Count > 0)
                {
                    throw CreateConversionException(
                        stage,
                        result,
                        $"Table '{tableName}' cannot be converted because preflight found blocking issues. " +
                        string.Join(" ", unhandledBlockers.Select(finding => finding.Message)));
                }

                stage = TableConversionFailureStage.Snapshot;
                ct.ThrowIfCancellationRequested();
                snapshot = RangeRollbackSnapshot.Create(ctx.Book, sheet, effectiveRange, ct);

                stage = TableConversionFailureStage.Normalization;
                if (mergedHeaderPolicy == TableMergedHeaderPolicy.UnmergeAndFill)
                {
                    sourceMutationStarted |= preflight.Findings.Any(
                        finding => finding.Kind == TablePreflightFindingKind.MergedCells
                            && finding.Addresses.Count > 0);
                    NormalizeMergedHeaders(
                        sheet,
                        effectiveRange,
                        preflight,
                        result.NormalizedMergedRanges,
                        ct);
                }

                if (headerPolicy == TableHeaderPolicy.Normalize)
                {
                    sourceMutationStarted |= preflight.Findings.Any(
                        finding => finding.Kind is TablePreflightFindingKind.BlankHeaders
                            or TablePreflightFindingKind.DuplicateHeaders);
                    NormalizeHeaders(
                        effectiveRange,
                        result.HeaderChanges,
                        ct);
                }

                ct.ThrowIfCancellationRequested();
                TablePreflightResult normalizedPreflight = AnalyzePreflight(
                    ctx.Book,
                    effectiveRange,
                    batch.WorkbookPath,
                    sheetName,
                    tableName,
                    rangeAddress,
                    hasHeaders,
                    ct);
                List<TablePreflightFinding> remainingBlockers = normalizedPreflight.Findings
                    .Where(finding => finding.Severity == TablePreflightSeverity.Blocker)
                    .ToList();
                if (remainingBlockers.Count > 0)
                {
                    throw new InvalidOperationException(
                        "Explicit normalization did not resolve all table creation blockers. " +
                        string.Join(" ", remainingBlockers.Select(finding => finding.Message)));
                }

                List<string> expectedHeaders = ReadHeaderValues(effectiveRange, ct);
                stage = TableConversionFailureStage.Creation;
                ct.ThrowIfCancellationRequested();
                listObjects = sheet.ListObjects;
                sourceMutationStarted = true;
                int headerOption = hasHeaders ? 1 : 0;
                newTable = listObjects.Add(1, effectiveRange, null, headerOption);
                newTable.Name = tableName;

                if (!string.IsNullOrWhiteSpace(tableStyle))
                {
                    stage = TableConversionFailureStage.Styling;
                    ct.ThrowIfCancellationRequested();
                    newTable.TableStyle = tableStyle;
                }

                stage = TableConversionFailureStage.Validation;
                ct.ThrowIfCancellationRequested();
                result.Validation = ValidateConvertedTable(
                    newTable,
                    effectiveRange,
                    snapshot,
                    tableName,
                    tableStyle,
                    hasHeaders,
                    expectedHeaders,
                    ct);
                if (!result.Validation.IsValid)
                {
                    throw new InvalidOperationException(
                        "The created table failed validation. " +
                        string.Join(" ", result.Validation.Findings.Select(finding => finding.Message)));
                }

                result.Table = ReadConvertedTableInfo(newTable, sheetName, tableName);

                stage = TableConversionFailureStage.Rollback;
                snapshot.DeleteBackup();
                result.Rollback = new TableRollbackResult();
                result.Success = true;
                return result;
            }
            catch (TableRangeConversionException)
            {
                throw;
            }
            catch (Exception ex)
            {
                var rollback = new TableRollbackResult
                {
                    Required = sourceMutationStarted,
                    Attempted = sourceMutationStarted
                };

                if (sourceMutationStarted && snapshot != null && sheet != null && effectiveRange != null)
                {
                    var rollbackErrors = new List<string>();
                    try
                    {
                        RemoveCreatedTable(sheet, effectiveRange, snapshot, ref newTable);
                    }
                    catch (Exception removeException)
                    {
                        rollbackErrors.Add(
                            $"Created table removal failed: {removeException.GetType().Name}: {removeException.Message}");
                    }

                    try
                    {
                        snapshot.Restore(effectiveRange);
                    }
                    catch (Exception restoreException)
                    {
                        rollbackErrors.Add(
                            $"Range restoration failed: {restoreException.GetType().Name}: {restoreException.Message}");
                    }

                    try
                    {
                        rollback.Verified = !TableExists(ctx.Book, tableName)
                            && snapshot.Verify(effectiveRange, ct: CancellationToken.None);
                        if (!rollback.Verified)
                        {
                            rollbackErrors.Add(
                                "The restored range did not match the captured values, formulas, representative formatting, or merged ranges.");
                        }
                    }
                    catch (Exception verificationException)
                    {
                        rollbackErrors.Add(
                            $"Rollback verification failed: {verificationException.GetType().Name}: {verificationException.Message}");
                    }

                    if (rollback.Verified)
                    {
                        try
                        {
                            snapshot.DeleteBackup();
                            rollback.Completed = true;
                        }
                        catch (Exception cleanupException)
                        {
                            rollback.Verified = false;
                            rollbackErrors.Add(
                                $"Rollback snapshot cleanup failed: {cleanupException.GetType().Name}: {cleanupException.Message}");
                        }
                    }

                    if (!rollback.Verified)
                    {
                        rollback.RecoverySheetName = snapshot.BackupSheetName;
                    }

                    rollback.ErrorMessage = rollbackErrors.Count == 0
                        ? null
                        : string.Join(" ", rollbackErrors);
                }
                else if (snapshot != null)
                {
                    try
                    {
                        snapshot.DeleteBackup();
                    }
                    catch (Exception cleanupException)
                    {
                        rollback.ErrorMessage =
                            $"Rollback snapshot cleanup failed: {cleanupException.GetType().Name}: {cleanupException.Message}";
                        rollback.RecoverySheetName = snapshot.BackupSheetName;
                    }
                }

                var failedStage = rollback.ErrorMessage != null
                    ? TableConversionFailureStage.Rollback
                    : stage;
                if (ex is OperationCanceledException cancellationException)
                {
                    throw CreateConversionException(
                        failedStage,
                        result,
                        BuildConversionFailureMessage(tableName, stage, cancellationException, rollback),
                        rollback,
                        cancellationException);
                }

                throw CreateConversionException(
                    failedStage,
                    result,
                    BuildConversionFailureMessage(tableName, stage, ex, rollback),
                    rollback,
                    ex);
            }
            finally
            {
                snapshot?.Dispose();
                ComUtilities.Release(ref newTable);
                ComUtilities.Release(ref listObjects);
                ComUtilities.Release(ref effectiveRange);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void ValidateConversionPolicies(
        bool hasHeaders,
        TableMergedHeaderPolicy mergedHeaderPolicy,
        TableHeaderPolicy headerPolicy)
    {
        if (!Enum.IsDefined(mergedHeaderPolicy))
        {
            throw new ArgumentOutOfRangeException(nameof(mergedHeaderPolicy));
        }

        if (!Enum.IsDefined(headerPolicy))
        {
            throw new ArgumentOutOfRangeException(nameof(headerPolicy));
        }

        if (!hasHeaders)
        {
            throw new ArgumentException(
                "convert-range requires hasHeaders=true because Excel inserts a new header row for headerless ranges. Use table create for headerless data.");
        }
    }

    private static List<TablePreflightFinding> GetUnhandledConversionBlockers(
        Excel.Worksheet sheet,
        Excel.Range effectiveRange,
        TablePreflightResult preflight,
        bool hasHeaders,
        TableMergedHeaderPolicy mergedHeaderPolicy,
        TableHeaderPolicy headerPolicy,
        CancellationToken ct)
    {
        var blockers = new List<TablePreflightFinding>();
        foreach (TablePreflightFinding finding in preflight.Findings)
        {
            ct.ThrowIfCancellationRequested();
            if (finding.Severity != TablePreflightSeverity.Blocker)
            {
                continue;
            }

            bool handled = finding.Kind switch
            {
                TablePreflightFindingKind.MergedCells =>
                    hasHeaders
                    && mergedHeaderPolicy == TableMergedHeaderPolicy.UnmergeAndFill
                    && finding.Addresses.All(address =>
                        IsContainedHeaderMerge(sheet, effectiveRange, address)),
                TablePreflightFindingKind.BlankHeaders
                    or TablePreflightFindingKind.DuplicateHeaders =>
                    hasHeaders && headerPolicy == TableHeaderPolicy.Normalize,
                _ => false
            };

            if (!handled)
            {
                blockers.Add(finding);
            }
        }

        return blockers;
    }

    private static void AddConversionOnlyBlockers(
        Excel.Range effectiveRange,
        TablePreflightResult preflight,
        CancellationToken ct)
    {
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        Excel.Range? headerRange = null;
        Excel.Range? headerCells = null;
        try
        {
            rows = effectiveRange.Rows;
            columns = effectiveRange.Columns;
            int rowCount = Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture);
            if (rowCount < 2)
            {
                preflight.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.HeaderOnlyRange,
                    Severity = TablePreflightSeverity.Blocker,
                    Addresses = [effectiveRange.Address],
                    Message = "The effective range contains only a header row, so Excel would insert a data row and shift following cells.",
                    Remediation = "Include at least one data row before converting the range."
                });
            }

            headerRange = rows[1];
            headerCells = headerRange.Cells;
            int columnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
            var lossyAddresses = new List<string>();
            for (int index = 1; index <= columnCount; index++)
            {
                ct.ThrowIfCancellationRequested();
                Excel.Range? cell = null;
                try
                {
                    cell = headerCells[1, index];
                    object? value = cell.Value2;
                    string formula = Convert.ToString(
                        cell.Formula2,
                        CultureInfo.InvariantCulture) ?? string.Empty;
                    string text = Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty;
                    if (formula.StartsWith('=')
                        || value is not null and not DBNull and not string
                        || text.Length > 255)
                    {
                        lossyAddresses.Add(cell.Address);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref cell);
                }
            }

            if (lossyAddresses.Count > 0)
            {
                preflight.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.LossyHeaders,
                    Severity = TablePreflightSeverity.Blocker,
                    Addresses = lossyAddresses,
                    Message = "One or more headers contain formulas, non-text values, or more than 255 characters that Excel would silently convert or truncate.",
                    Remediation = "Replace the listed headers with text values of at most 255 characters before converting."
                });
            }
        }
        finally
        {
            ComUtilities.Release(ref headerCells);
            ComUtilities.Release(ref headerRange);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
        }
    }

    private static bool IsContainedHeaderMerge(
        Excel.Worksheet sheet,
        Excel.Range effectiveRange,
        string address)
    {
        Excel.Range? mergedRange = null;
        Excel.Range? mergedRows = null;
        Excel.Range? mergedColumns = null;
        Excel.Range? effectiveRows = null;
        Excel.Range? effectiveColumns = null;
        try
        {
            mergedRange = sheet.Range[address];
            mergedRows = mergedRange.Rows;
            mergedColumns = mergedRange.Columns;
            effectiveRows = effectiveRange.Rows;
            effectiveColumns = effectiveRange.Columns;

            int mergedRow = Convert.ToInt32(mergedRange.Row, CultureInfo.InvariantCulture);
            int mergedColumn = Convert.ToInt32(mergedRange.Column, CultureInfo.InvariantCulture);
            int mergedRowCount = Convert.ToInt32(mergedRows.Count, CultureInfo.InvariantCulture);
            int mergedColumnCount = Convert.ToInt32(mergedColumns.Count, CultureInfo.InvariantCulture);
            int effectiveRow = Convert.ToInt32(effectiveRange.Row, CultureInfo.InvariantCulture);
            int effectiveColumn = Convert.ToInt32(effectiveRange.Column, CultureInfo.InvariantCulture);
            int effectiveRowCount = Convert.ToInt32(effectiveRows.Count, CultureInfo.InvariantCulture);
            int effectiveColumnCount = Convert.ToInt32(effectiveColumns.Count, CultureInfo.InvariantCulture);

            return mergedRow == effectiveRow
                && mergedRowCount == 1
                && mergedColumn >= effectiveColumn
                && mergedColumn + mergedColumnCount <= effectiveColumn + effectiveColumnCount
                && mergedRow < effectiveRow + effectiveRowCount;
        }
        finally
        {
            ComUtilities.Release(ref effectiveColumns);
            ComUtilities.Release(ref effectiveRows);
            ComUtilities.Release(ref mergedColumns);
            ComUtilities.Release(ref mergedRows);
            ComUtilities.Release(ref mergedRange);
        }
    }

    private static void NormalizeMergedHeaders(
        Excel.Worksheet sheet,
        Excel.Range effectiveRange,
        TablePreflightResult preflight,
        List<string> normalizedRanges,
        CancellationToken ct)
    {
        TablePreflightFinding? mergedFinding = preflight.Findings.FirstOrDefault(
            finding => finding.Kind == TablePreflightFindingKind.MergedCells);
        if (mergedFinding == null)
        {
            return;
        }

        foreach (string address in mergedFinding.Addresses)
        {
            ct.ThrowIfCancellationRequested();
            if (!IsContainedHeaderMerge(sheet, effectiveRange, address))
            {
                throw new InvalidOperationException(
                    $"Merged range '{address}' is not wholly contained in the table header row.");
            }

            Excel.Range? mergedRange = null;
            Excel.Range? mergedCells = null;
            Excel.Range? firstCell = null;
            try
            {
                mergedRange = sheet.Range[address];
                mergedCells = mergedRange.Cells;
                firstCell = mergedCells[1, 1];
                object? value = firstCell.Value2;
                mergedRange.UnMerge();
                mergedRange.Value2 = value;
                normalizedRanges.Add(address);
            }
            finally
            {
                ComUtilities.Release(ref firstCell);
                ComUtilities.Release(ref mergedCells);
                ComUtilities.Release(ref mergedRange);
            }
        }

    }

    private static void NormalizeHeaders(
        Excel.Range effectiveRange,
        List<TableHeaderChange> changes,
        CancellationToken ct)
    {
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        Excel.Range? headerRange = null;
        Excel.Range? headerCells = null;
        try
        {
            rows = effectiveRange.Rows;
            columns = effectiveRange.Columns;
            headerRange = rows[1];
            headerCells = headerRange.Cells;
            object values = headerRange.Value2;
            int count = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
            int row = Convert.ToInt32(effectiveRange.Row, CultureInfo.InvariantCulture);
            int firstColumn = Convert.ToInt32(effectiveRange.Column, CultureInfo.InvariantCulture);
            var used = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int offset = 0; offset < count; offset++)
            {
                ct.ThrowIfCancellationRequested();
                object? rawValue = GetMatrixValue(values, 0, offset);
                string? original = Convert.ToString(rawValue, CultureInfo.InvariantCulture);
                string trimmed = original?.Trim() ?? string.Empty;
                TableHeaderChangeReason? reason = null;
                string candidate = trimmed;

                if (trimmed.Length == 0)
                {
                    reason = TableHeaderChangeReason.Blank;
                    candidate = $"Column{offset + 1}";
                }
                else if (used.Contains(trimmed))
                {
                    reason = TableHeaderChangeReason.Duplicate;
                    candidate = trimmed;
                }

                string unique = EnsureUniqueHeader(candidate, used);
                if (reason.HasValue || !string.Equals(unique, candidate, StringComparison.Ordinal))
                {
                    reason ??= TableHeaderChangeReason.Duplicate;
                    Excel.Range? cell = null;
                    try
                    {
                        cell = headerCells[1, offset + 1];
                        cell.Value2 = unique;
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                    }

                    changes.Add(new TableHeaderChange
                    {
                        Address = GetAbsoluteAddress(firstColumn + offset, row),
                        OriginalValue = original,
                        NewValue = unique,
                        Reason = reason.Value
                    });
                }

                used.Add(unique);
            }

        }
        finally
        {
            ComUtilities.Release(ref headerCells);
            ComUtilities.Release(ref headerRange);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
        }
    }

    private static string EnsureUniqueHeader(string candidate, HashSet<string> used)
    {
        if (!used.Contains(candidate))
        {
            return candidate;
        }

        for (int suffix = 2; ; suffix++)
        {
            string suffixText = $"_{suffix}";
            int maximumBaseLength = 255 - suffixText.Length;
            string baseName = candidate.Length > maximumBaseLength
                ? candidate[..maximumBaseLength]
                : candidate;
            string unique = baseName + suffixText;
            if (!used.Contains(unique))
            {
                return unique;
            }
        }
    }

    private static List<string> ReadHeaderValues(
        Excel.Range effectiveRange,
        CancellationToken ct)
    {
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        Excel.Range? headerRange = null;
        try
        {
            rows = effectiveRange.Rows;
            columns = effectiveRange.Columns;
            headerRange = rows[1];
            object values = headerRange.Value2;
            int count = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
            var headers = new List<string>(count);
            for (int offset = 0; offset < count; offset++)
            {
                ct.ThrowIfCancellationRequested();
                headers.Add(
                    Convert.ToString(
                        GetMatrixValue(values, 0, offset),
                        CultureInfo.InvariantCulture) ?? string.Empty);
            }

            return headers;
        }
        finally
        {
            ComUtilities.Release(ref headerRange);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
        }
    }

    private static TableConversionValidationResult ValidateConvertedTable(
        dynamic table,
        Excel.Range effectiveRange,
        RangeRollbackSnapshot snapshot,
        string expectedName,
        string? expectedStyle,
        bool hasHeaders,
        IReadOnlyList<string> expectedHeaders,
        CancellationToken ct)
    {
        var validation = new TableConversionValidationResult();
        Excel.Range? tableRange = null;
        Excel.Range? dataBodyRange = null;
        dynamic? tableStyle = null;
        try
        {
            tableRange = table.Range;
            string actualRange = tableRange.Address;
            string actualName = table.Name;
            if (!string.Equals(actualName, expectedName, StringComparison.Ordinal)
                || !string.Equals(actualRange, effectiveRange.Address, StringComparison.OrdinalIgnoreCase))
            {
                validation.Findings.Add(new TableConversionValidationFinding
                {
                    Kind = TableConversionValidationFindingKind.TableMismatch,
                    Addresses = [actualRange],
                    Message = $"Created table identity or range did not match '{expectedName}' at '{effectiveRange.Address}'."
                });
            }

            if (!string.IsNullOrWhiteSpace(expectedStyle))
            {
                tableStyle = table.TableStyle;
                string actualStyle = tableStyle?.Name?.ToString() ?? string.Empty;
                if (!string.Equals(actualStyle, expectedStyle, StringComparison.OrdinalIgnoreCase))
                {
                    validation.Findings.Add(new TableConversionValidationFinding
                    {
                        Kind = TableConversionValidationFindingKind.StyleMismatch,
                        Addresses = [actualRange],
                        Message = $"Table style '{expectedStyle}' was not applied."
                    });
                }

            }

            List<string> actualHeaders = ReadTableColumnNames(table);
            if (!actualHeaders.SequenceEqual(expectedHeaders, StringComparer.Ordinal))
            {
                validation.Findings.Add(new TableConversionValidationFinding
                {
                    Kind = TableConversionValidationFindingKind.TableMismatch,
                    Addresses = [actualRange],
                    Message = "Excel changed one or more table headers during conversion."
                });
            }

            validation.ShowTotals = table.ShowTotals;
            if (validation.ShowTotals)
            {
                validation.Findings.Add(new TableConversionValidationFinding
                {
                    Kind = TableConversionValidationFindingKind.UnexpectedTotalsRow,
                    Addresses = [actualRange],
                    Message = "Excel unexpectedly enabled a totals row during range conversion."
                });
            }

            if (!snapshot.SourceContentMatches(effectiveRange, hasHeaders, ct))
            {
                validation.Findings.Add(new TableConversionValidationFinding
                {
                    Kind = TableConversionValidationFindingKind.SourceContentChanged,
                    Addresses = [effectiveRange.Address],
                    Message = "Non-header source values or formulas changed during table creation."
                });
            }

            dataBodyRange = table.DataBodyRange;
            if (dataBodyRange != null)
            {
                InspectCalculatedColumns(table, dataBodyRange, validation, ct);
            }

            validation.IsValid = validation.Findings.Count == 0;
            return validation;
        }
        finally
        {
            ComUtilities.Release(ref tableStyle);
            ComUtilities.Release(ref dataBodyRange);
            ComUtilities.Release(ref tableRange);
        }
    }

    private static void InspectCalculatedColumns(
        dynamic table,
        Excel.Range dataBodyRange,
        TableConversionValidationResult validation,
        CancellationToken ct)
    {
        Excel.Range? rows = null;
        Excel.Range? columnsRange = null;
        dynamic? listColumns = null;
        try
        {
            rows = dataBodyRange.Rows;
            columnsRange = dataBodyRange.Columns;
            listColumns = table.ListColumns;
            int rowCount = Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture);
            int columnCount = Convert.ToInt32(columnsRange.Count, CultureInfo.InvariantCulture);
            int firstRow = Convert.ToInt32(dataBodyRange.Row, CultureInfo.InvariantCulture);
            int firstColumn = Convert.ToInt32(dataBodyRange.Column, CultureInfo.InvariantCulture);
            object formulas = dataBodyRange.FormulaR1C1;
            dataBodyRange.Calculate();
            AddFormulaErrorFindings(dataBodyRange, validation, ct);

            for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
            {
                string? pattern = null;
                var formulaAddresses = new List<string>();
                var inconsistentAddresses = new List<string>();

                for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
                {
                    ct.ThrowIfCancellationRequested();
                    object? formulaValue = GetMatrixValue(formulas, rowOffset, columnOffset);
                    string formula = Convert.ToString(formulaValue, CultureInfo.InvariantCulture) ?? string.Empty;
                    string address = GetAbsoluteAddress(firstColumn + columnOffset, firstRow + rowOffset);
                    if (!formula.StartsWith('='))
                    {
                        if (pattern != null)
                        {
                            inconsistentAddresses.Add(address);
                        }

                        continue;
                    }

                    validation.FormulaCellsChecked++;
                    formulaAddresses.Add(address);
                    if (pattern == null)
                    {
                        pattern = formula;
                    }
                    else if (!string.Equals(pattern, formula, StringComparison.Ordinal))
                    {
                        inconsistentAddresses.Add(address);
                    }
                }

                if (formulaAddresses.Count == 0)
                {
                    continue;
                }

                if (formulaAddresses.Count != rowCount)
                {
                    for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
                    {
                        string formula = Convert.ToString(
                            GetMatrixValue(formulas, rowOffset, columnOffset),
                            CultureInfo.InvariantCulture) ?? string.Empty;
                        if (!formula.StartsWith('='))
                        {
                            inconsistentAddresses.Add(
                                GetAbsoluteAddress(firstColumn + columnOffset, firstRow + rowOffset));
                        }
                    }
                }

                if (inconsistentAddresses.Count > 0)
                {
                    validation.Findings.Add(new TableConversionValidationFinding
                    {
                        Kind = TableConversionValidationFindingKind.InconsistentCalculatedColumn,
                        Addresses = inconsistentAddresses.Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                        Message = "A formula-bearing table column is not a consistent calculated column."
                    });
                    continue;
                }

                dynamic? column = null;
                try
                {
                    column = listColumns.Item(columnOffset + 1);
                    validation.CalculatedColumns.Add(column.Name?.ToString() ?? string.Empty);
                }
                finally
                {
                    ComUtilities.Release(ref column);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref listColumns);
            ComUtilities.Release(ref columnsRange);
            ComUtilities.Release(ref rows);
        }
    }

    private static void AddFormulaErrorFindings(
        Excel.Range dataBodyRange,
        TableConversionValidationResult validation,
        CancellationToken ct)
    {
        const int ExcelNoCellsFound = unchecked((int)0x800A03EC);
        Excel.Range? errorCells = null;
        try
        {
            errorCells = dataBodyRange.SpecialCells(
                Excel.XlCellType.xlCellTypeFormulas,
                Excel.XlSpecialCellsValue.xlErrors);
            string address = errorCells.Address;
            validation.Findings.Add(new TableConversionValidationFinding
            {
                Kind = TableConversionValidationFindingKind.FormulaError,
                Addresses = [address],
                Message = $"Formula cells '{address}' evaluate to Excel errors."
            });
        }
        catch (COMException ex) when (ex.HResult == ExcelNoCellsFound)
        {
            List<string> addresses = FindFormulaErrorsByCell(dataBodyRange, ct);
            if (addresses.Count > 0)
            {
                validation.Findings.Add(new TableConversionValidationFinding
                {
                    Kind = TableConversionValidationFindingKind.FormulaError,
                    Addresses = addresses,
                    Message = $"Formula cells '{string.Join(", ", addresses)}' evaluate to Excel errors."
                });
            }
        }
        finally
        {
            ComUtilities.Release(ref errorCells);
        }
    }

    private static List<string> FindFormulaErrorsByCell(
        Excel.Range dataBodyRange,
        CancellationToken ct)
    {
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        Excel.Range? cells = null;
        var addresses = new List<string>();
        try
        {
            rows = dataBodyRange.Rows;
            columns = dataBodyRange.Columns;
            cells = dataBodyRange.Cells;
            int rowCount = Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture);
            int columnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
            for (int row = 1; row <= rowCount; row++)
            {
                for (int column = 1; column <= columnCount; column++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Range? cell = null;
                    try
                    {
                        cell = cells[row, column];
                        string formula = Convert.ToString(
                            cell.Formula2,
                            CultureInfo.InvariantCulture) ?? string.Empty;
                        if (!formula.StartsWith('='))
                        {
                            continue;
                        }

                        object? value = cell.Value2;
                        string text = Convert.ToString(cell.Text, CultureInfo.InvariantCulture) ?? string.Empty;
                        if (value is ErrorWrapper
                            || value is not string && IsExcelErrorText(text))
                        {
                            addresses.Add(cell.Address);
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref cell);
                    }
                }
            }

            return addresses;
        }
        finally
        {
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
        }
    }

    private static bool IsExcelErrorText(string text)
    {
        string[] errors =
        [
            "#NULL!",
            "#DIV/0!",
            "#VALUE!",
            "#REF!",
            "#NAME?",
            "#NUM!",
            "#N/A",
            "#GETTING_DATA",
            "#SPILL!",
            "#CALC!",
            "#FIELD!",
            "#BLOCKED!",
            "#UNKNOWN!",
            "#CONNECT!",
            "#BUSY!"
        ];
        return errors.Any(error => text.StartsWith(error, StringComparison.OrdinalIgnoreCase));
    }

    private static List<string> ReadTableColumnNames(dynamic table)
    {
        dynamic? columns = null;
        try
        {
            columns = table.ListColumns;
            var names = new List<string>();
            for (int index = 1; index <= columns.Count; index++)
            {
                dynamic? column = null;
                try
                {
                    column = columns.Item(index);
                    names.Add(column.Name?.ToString() ?? string.Empty);
                }
                finally
                {
                    ComUtilities.Release(ref column);
                }
            }

            return names;
        }
        finally
        {
            ComUtilities.Release(ref columns);
        }
    }

    private static TableInfo ReadConvertedTableInfo(dynamic table, string sheetName, string tableName)
    {
        Excel.Range? tableRange = null;
        Excel.Range? dataBodyRange = null;
        Excel.Range? dataBodyRows = null;
        dynamic? columns = null;
        dynamic? tableStyle = null;
        try
        {
            tableRange = table.Range;
            dataBodyRange = table.DataBodyRange;
            dataBodyRows = dataBodyRange?.Rows;
            columns = table.ListColumns;
            tableStyle = table.TableStyle;
            var names = new List<string>();
            for (int index = 1; index <= columns.Count; index++)
            {
                dynamic? column = null;
                try
                {
                    column = columns.Item(index);
                    names.Add(column.Name?.ToString() ?? string.Empty);
                }
                finally
                {
                    ComUtilities.Release(ref column);
                }
            }

            return new TableInfo
            {
                Name = tableName,
                SheetName = sheetName,
                Range = tableRange.Address,
                HasHeaders = table.ShowHeaders,
                TableStyle = tableStyle?.Name?.ToString() ?? string.Empty,
                RowCount = dataBodyRows == null ? 0 : Convert.ToInt32(dataBodyRows.Count, CultureInfo.InvariantCulture),
                ColumnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture),
                Columns = names,
                ShowTotals = table.ShowTotals
            };
        }
        finally
        {
            ComUtilities.Release(ref tableStyle);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref dataBodyRows);
            ComUtilities.Release(ref dataBodyRange);
            ComUtilities.Release(ref tableRange);
        }
    }

    private static void RemoveCreatedTable(
        Excel.Worksheet sheet,
        Excel.Range effectiveRange,
        RangeRollbackSnapshot snapshot,
        ref dynamic? table)
    {
        if (table != null)
        {
            table.Unlist();
            ComUtilities.Release(ref table);
            return;
        }

        dynamic? listObjects = null;
        try
        {
            listObjects = sheet.ListObjects;
            for (int index = listObjects.Count; index >= 1; index--)
            {
                dynamic? candidate = null;
                Excel.Range? candidateRange = null;
                try
                {
                    candidate = listObjects.Item(index);
                    string candidateName = candidate.Name?.ToString() ?? string.Empty;
                    candidateRange = candidate.Range;
                    if (!snapshot.ExistingTableNames.Contains(candidateName)
                        && string.Equals(
                            candidateRange.Address,
                            effectiveRange.Address,
                            StringComparison.OrdinalIgnoreCase))
                    {
                        candidate.Unlist();
                        return;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref candidateRange);
                    ComUtilities.Release(ref candidate);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref listObjects);
        }
    }

    private static TableRangeConversionException CreateConversionException(
        TableConversionFailureStage stage,
        TableRangeConversionResult result,
        string message,
        TableRollbackResult? rollback = null,
        Exception? innerException = null)
    {
        return new TableRangeConversionException(
            message,
            new TableRangeConversionFailureDetails
            {
                FailureStage = stage,
                WasCancelled = innerException is OperationCanceledException
                    && BatchExecutionCancellation.Current.IsCancellationRequested,
                WasTimedOut = innerException is OperationCanceledException
                    && !BatchExecutionCancellation.Current.IsCancellationRequested,
                SheetName = result.SheetName,
                TableName = result.TableName,
                RequestedRange = result.RequestedRange,
                EffectiveRange = string.IsNullOrEmpty(result.EffectiveRange) ? null : result.EffectiveRange,
                PreflightFindings = result.PreflightFindings,
                HeaderChanges = result.HeaderChanges,
                Validation = result.Validation,
                Rollback = rollback ?? result.Rollback
            },
            innerException);
    }

    private static string BuildConversionFailureMessage(
        string tableName,
        TableConversionFailureStage stage,
        Exception exception,
        TableRollbackResult rollback)
    {
        string rollbackStatus = !rollback.Required
            ? "No source rollback was required."
            : rollback.Verified
                ? "The original range state was restored and verified."
                : $"The original range state could not be fully verified after rollback. {rollback.ErrorMessage}";
        return $"Table '{tableName}' conversion failed during {stage}. " +
            $"{exception.GetType().Name}: {exception.Message} {rollbackStatus}";
    }

    private sealed class RangeRollbackSnapshot : IDisposable
    {
        private readonly Excel.Worksheet _sourceSheet;
        private readonly object? _formulaState;
        private readonly object? _valueState;
        private readonly object? _numberFormatState;
        private readonly List<string> _mergedRanges;
        private Excel.Worksheet? _backupSheet;
        private Excel.Range? _backupRange;
        private bool _deleted;

        private RangeRollbackSnapshot(
            Excel.Worksheet sourceSheet,
            object? formulaState,
            object? valueState,
            object? numberFormatState,
            List<string> mergedRanges,
            HashSet<string> existingTableNames,
            Excel.Worksheet backupSheet,
            Excel.Range backupRange)
        {
            _sourceSheet = sourceSheet;
            _formulaState = formulaState;
            _valueState = valueState;
            _numberFormatState = numberFormatState;
            _mergedRanges = mergedRanges;
            ExistingTableNames = existingTableNames;
            _backupSheet = backupSheet;
            _backupRange = backupRange;
        }

        internal HashSet<string> ExistingTableNames { get; }

        internal string? BackupSheetName => _backupSheet?.Name;

        internal static RangeRollbackSnapshot Create(
            Excel.Workbook workbook,
            Excel.Worksheet sourceSheet,
            Excel.Range sourceRange,
            CancellationToken ct)
        {
            Excel.Sheets? worksheets = null;
            Excel.Worksheet? afterSheet = null;
            Excel.Worksheet? backupSheet = null;
            Excel.Range? backupRange = null;
            try
            {
                ct.ThrowIfCancellationRequested();
                object? formulas = sourceRange.Formula2;
                object? values = sourceRange.Value2;
                object? numberFormats = sourceRange.NumberFormat;
                List<string> merges = RangeMergeDiscovery.CollectMergedRanges(sourceRange, ct);
                HashSet<string> existingTables = CollectTableNames(workbook, ct);

                worksheets = workbook.Worksheets;
                afterSheet = (Excel.Worksheet)worksheets[worksheets.Count];
                backupSheet = (Excel.Worksheet)worksheets.Add(After: afterSheet);
                backupSheet.Name = CreateUniqueRollbackSheetName(workbook);
                backupSheet.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;
                backupRange = backupSheet.Range[sourceRange.Address];
                sourceRange.Copy(backupRange);
                backupRange.Formula2 = formulas;
                var snapshot = new RangeRollbackSnapshot(
                    sourceSheet,
                    formulas,
                    values,
                    numberFormats,
                    merges,
                    existingTables,
                    backupSheet,
                    backupRange);
                backupSheet = null;
                backupRange = null;
                return snapshot;
            }
            catch (Exception snapshotException)
            {
                if (backupSheet != null)
                {
                    try
                    {
                        TryDeleteSheet(backupSheet);
                    }
                    catch (Exception cleanupException)
                    {
                        throw new AggregateException(
                            "Rollback snapshot creation and cleanup both failed.",
                            snapshotException,
                            cleanupException);
                    }
                }

                throw;
            }
            finally
            {
                ComUtilities.Release(ref backupRange);
                ComUtilities.Release(ref backupSheet);
                ComUtilities.Release(ref afterSheet);
                ComUtilities.Release(ref worksheets);
            }
        }

        internal void Restore(Excel.Range targetRange)
        {
            if (_backupRange == null)
            {
                throw new InvalidOperationException("The rollback snapshot is no longer available.");
            }

            targetRange.UnMerge();
            targetRange.Clear();
            _backupRange.Copy(targetRange);
            targetRange.UnMerge();
            targetRange.Formula2 = _formulaState;

            foreach (string address in _mergedRanges)
            {
                Excel.Range? mergedRange = null;
                try
                {
                    mergedRange = _sourceSheet.Range[address];
                    mergedRange.Merge();
                }
                finally
                {
                    ComUtilities.Release(ref mergedRange);
                }
            }
        }

        internal bool Verify(Excel.Range targetRange, CancellationToken ct)
        {
            ct.ThrowIfCancellationRequested();
            Excel.Range? backupRange = _backupRange;
            if (backupRange == null)
            {
                return false;
            }

            return ContentStateEquals(targetRange)
                && StateEquals(_numberFormatState, targetRange.NumberFormat)
                && FormatsEqual(backupRange, targetRange, ct)
                && MergedRangesEqual(targetRange, ct);
        }

        internal bool SourceContentMatches(
            Excel.Range currentRange,
            bool hasHeaders,
            CancellationToken ct)
        {
            Excel.Range? rows = null;
            Excel.Range? dataRange = null;
            Excel.Range? dataColumns = null;
            try
            {
                rows = currentRange.Rows;
                int rowCount = Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture);
                int firstDataRow = hasHeaders ? 2 : 1;
                if (rowCount < firstDataRow)
                {
                    return true;
                }

                dataRange = rows[$"{firstDataRow}:{rowCount}"];
                dataColumns = dataRange.Columns;
                object currentFormulas = dataRange.Formula2;
                int columnCount = Convert.ToInt32(dataColumns.Count, CultureInfo.InvariantCulture);
                int dataRowCount = rowCount - firstDataRow + 1;

                for (int rowOffset = 0; rowOffset < dataRowCount; rowOffset++)
                {
                    for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
                    {
                        ct.ThrowIfCancellationRequested();
                        int snapshotRowOffset = rowOffset + firstDataRow - 1;
                        object? originalFormula = GetMatrixValue(_formulaState!, snapshotRowOffset, columnOffset);
                        object? currentFormula = GetMatrixValue(currentFormulas, rowOffset, columnOffset);
                        bool originallyFormula = Convert.ToString(
                            originalFormula,
                            CultureInfo.InvariantCulture)?.StartsWith('=') == true;
                        if (originallyFormula)
                        {
                            if (Convert.ToString(currentFormula, CultureInfo.InvariantCulture)?.StartsWith('=') != true)
                            {
                                return false;
                            }
                        }
                        else if (!ScalarEquals(originalFormula, currentFormula))
                        {
                            return false;
                        }
                    }
                }

                return true;
            }
            finally
            {
                ComUtilities.Release(ref dataColumns);
                ComUtilities.Release(ref dataRange);
                ComUtilities.Release(ref rows);
            }
        }

        private bool ContentStateEquals(Excel.Range targetRange)
        {
            Excel.Range? rows = null;
            Excel.Range? columns = null;
            try
            {
                rows = targetRange.Rows;
                columns = targetRange.Columns;
                int rowCount = Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture);
                int columnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
                object currentFormulas = targetRange.Formula2;
                object currentValues = targetRange.Value2;

                for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
                {
                    for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
                    {
                        object? originalFormula = GetMatrixValue(
                            _formulaState!,
                            rowOffset,
                            columnOffset);
                        object? currentFormula = GetMatrixValue(
                            currentFormulas,
                            rowOffset,
                            columnOffset);
                        bool isFormula = Convert.ToString(
                            originalFormula,
                            CultureInfo.InvariantCulture)?.StartsWith('=') == true;
                        if (isFormula)
                        {
                            if (!ScalarEquals(originalFormula, currentFormula))
                            {
                                return false;
                            }
                        }
                        else if (!ScalarEquals(
                            GetMatrixValue(_valueState!, rowOffset, columnOffset),
                            GetMatrixValue(currentValues, rowOffset, columnOffset)))
                        {
                            return false;
                        }
                    }
                }

                return true;
            }
            finally
            {
                ComUtilities.Release(ref columns);
                ComUtilities.Release(ref rows);
            }
        }

        internal void DeleteBackup()
        {
            if (_deleted)
            {
                return;
            }

            if (_backupSheet != null)
            {
                TryDeleteSheet(_backupSheet);
            }

            ComUtilities.Release(ref _backupRange);
            ComUtilities.Release(ref _backupSheet);
            _deleted = true;
        }

        public void Dispose()
        {
            ComUtilities.Release(ref _backupRange);
            ComUtilities.Release(ref _backupSheet);
        }

        private bool MergedRangesEqual(Excel.Range targetRange, CancellationToken ct)
        {
            List<string> restored = RangeMergeDiscovery.CollectMergedRanges(targetRange, ct);
            return _mergedRanges.OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .SequenceEqual(
                    restored.OrderBy(value => value, StringComparer.OrdinalIgnoreCase),
                    StringComparer.OrdinalIgnoreCase);
        }

        private static bool FormatsEqual(
            Excel.Range expectedRange,
            Excel.Range actualRange,
            CancellationToken ct)
        {
            Excel.Range? expectedRows = null;
            Excel.Range? expectedColumns = null;
            Excel.Range? expectedCells = null;
            Excel.Range? actualRows = null;
            Excel.Range? actualColumns = null;
            Excel.Range? actualCells = null;
            try
            {
                expectedRows = expectedRange.Rows;
                expectedColumns = expectedRange.Columns;
                expectedCells = expectedRange.Cells;
                actualRows = actualRange.Rows;
                actualColumns = actualRange.Columns;
                actualCells = actualRange.Cells;
                int rowCount = Convert.ToInt32(expectedRows.Count, CultureInfo.InvariantCulture);
                int columnCount = Convert.ToInt32(expectedColumns.Count, CultureInfo.InvariantCulture);
                if (rowCount != Convert.ToInt32(actualRows.Count, CultureInfo.InvariantCulture)
                    || columnCount != Convert.ToInt32(actualColumns.Count, CultureInfo.InvariantCulture))
                {
                    return false;
                }

                for (int row = 1; row <= rowCount; row++)
                {
                    for (int column = 1; column <= columnCount; column++)
                    {
                        ct.ThrowIfCancellationRequested();
                        if (CaptureCellFormat(expectedCells, row, column)
                            != CaptureCellFormat(actualCells, row, column))
                        {
                            return false;
                        }
                    }
                }

                return true;
            }
            finally
            {
                ComUtilities.Release(ref actualCells);
                ComUtilities.Release(ref actualColumns);
                ComUtilities.Release(ref actualRows);
                ComUtilities.Release(ref expectedCells);
                ComUtilities.Release(ref expectedColumns);
                ComUtilities.Release(ref expectedRows);
            }
        }

        private static CellFormatSample CaptureCellFormat(
            Excel.Range cells,
            int row,
            int column)
        {
            Excel.Range? cell = null;
            Excel.Font? font = null;
            Excel.Interior? interior = null;
            Excel.Borders? borders = null;
            try
            {
                cell = cells[row, column];
                font = cell.Font;
                interior = cell.Interior;
                borders = cell.Borders;
                return new CellFormatSample(
                    row,
                    column,
                    FormatValue(cell.NumberFormat),
                    FormatValue(cell.Style),
                    FormatValue(cell.HorizontalAlignment),
                    FormatValue(cell.VerticalAlignment),
                    FormatValue(cell.WrapText),
                    FormatValue(font.Name),
                    FormatValue(font.Size),
                    FormatValue(font.Bold),
                    FormatValue(font.Italic),
                    FormatValue(font.Color),
                    FormatValue(interior.Color),
                    FormatValue(interior.Pattern),
                    CaptureBorder(borders, Excel.XlBordersIndex.xlEdgeLeft),
                    CaptureBorder(borders, Excel.XlBordersIndex.xlEdgeTop),
                    CaptureBorder(borders, Excel.XlBordersIndex.xlEdgeRight),
                    CaptureBorder(borders, Excel.XlBordersIndex.xlEdgeBottom));
            }
            finally
            {
                ComUtilities.Release(ref borders);
                ComUtilities.Release(ref interior);
                ComUtilities.Release(ref font);
                ComUtilities.Release(ref cell);
            }
        }

        private static BorderFormatSample CaptureBorder(
            Excel.Borders borders,
            Excel.XlBordersIndex index)
        {
            Excel.Border? border = null;
            try
            {
                border = borders[index];
                return new BorderFormatSample(
                    FormatValue(border.LineStyle),
                    FormatValue(border.Weight),
                    FormatValue(border.Color));
            }
            finally
            {
                ComUtilities.Release(ref border);
            }
        }

        private static string FormatValue(object? value)
        {
            if (value is null or DBNull)
            {
                return "<null>";
            }

            return Convert.ToString(value, CultureInfo.InvariantCulture) ?? "<null>";
        }

        private static HashSet<string> CollectTableNames(Excel.Workbook workbook, CancellationToken ct)
        {
            Excel.Sheets? sheets = null;
            var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                sheets = workbook.Worksheets;
                for (int sheetIndex = 1; sheetIndex <= sheets.Count; sheetIndex++)
                {
                    ct.ThrowIfCancellationRequested();
                    Excel.Worksheet? sheet = null;
                    dynamic? tables = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets[sheetIndex];
                        tables = sheet.ListObjects;
                        for (int tableIndex = 1; tableIndex <= tables.Count; tableIndex++)
                        {
                            dynamic? table = null;
                            try
                            {
                                table = tables.Item(tableIndex);
                                names.Add(table.Name?.ToString() ?? string.Empty);
                            }
                            finally
                            {
                                ComUtilities.Release(ref table);
                            }
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref tables);
                        ComUtilities.Release(ref sheet);
                    }
                }

                return names;
            }
            finally
            {
                ComUtilities.Release(ref sheets);
            }
        }

        private static string CreateUniqueRollbackSheetName(Excel.Workbook workbook)
        {
            for (int attempt = 0; attempt < 10; attempt++)
            {
                string candidate = RollbackSheetPrefix + Guid.NewGuid().ToString("N")[..12];
                Excel.Worksheet? existing = null;
                try
                {
                    existing = ComUtilities.FindSheet(workbook, candidate);
                    if (existing == null)
                    {
                        return candidate;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref existing);
                }
            }

            throw new InvalidOperationException("Could not allocate a unique rollback worksheet name.");
        }

        private static void TryDeleteSheet(Excel.Worksheet sheet)
        {
            Excel.Application? application = null;
            Excel.XlSheetVisibility originalVisibility = Excel.XlSheetVisibility.xlSheetVisible;
            bool visibilityChanged = false;
            try
            {
                application = sheet.Application;
                bool originalAlerts = application.DisplayAlerts;
                application.DisplayAlerts = false;
                try
                {
                    originalVisibility = sheet.Visible;
                    sheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
                    visibilityChanged = true;
                    sheet.Delete();
                    visibilityChanged = false;
                }
                finally
                {
                    application.DisplayAlerts = originalAlerts;
                }
            }
            catch
            {
                if (visibilityChanged)
                {
                    sheet.Visible = originalVisibility;
                }

                throw;
            }
            finally
            {
                ComUtilities.Release(ref application);
            }
        }

        private static bool StateEquals(object? expected, object? actual)
        {
            if (expected is Array || actual is Array)
            {
                if (expected is not Array left || actual is not Array right
                    || left.Rank != right.Rank
                    || left.Length != right.Length)
                {
                    return false;
                }

                int rows = left.GetLength(0);
                int columns = left.Rank == 1 ? 1 : left.GetLength(1);
                for (int row = 0; row < rows; row++)
                {
                    for (int column = 0; column < columns; column++)
                    {
                        object? leftValue = left.Rank == 1
                            ? left.GetValue(left.GetLowerBound(0) + row)
                            : left.GetValue(left.GetLowerBound(0) + row, left.GetLowerBound(1) + column);
                        object? rightValue = right.Rank == 1
                            ? right.GetValue(right.GetLowerBound(0) + row)
                            : right.GetValue(right.GetLowerBound(0) + row, right.GetLowerBound(1) + column);
                        if (!ScalarEquals(leftValue, rightValue))
                        {
                            return false;
                        }
                    }
                }

                return true;
            }

            return ScalarEquals(expected, actual);
        }

        private static bool ScalarEquals(object? expected, object? actual)
        {
            if (expected is null or DBNull && actual is null or DBNull)
            {
                return true;
            }

            if (expected is ErrorWrapper expectedError && actual is ErrorWrapper actualError)
            {
                return expectedError.ErrorCode == actualError.ErrorCode;
            }

            return Equals(expected, actual);
        }

        private sealed record CellFormatSample(
            int Row,
            int Column,
            string NumberFormat,
            string Style,
            string HorizontalAlignment,
            string VerticalAlignment,
            string WrapText,
            string FontName,
            string FontSize,
            string FontBold,
            string FontItalic,
            string FontColor,
            string InteriorColor,
            string InteriorPattern,
            BorderFormatSample LeftBorder,
            BorderFormatSample TopBorder,
            BorderFormatSample RightBorder,
            BorderFormatSample BottomBorder);

        private sealed record BorderFormatSample(
            string LineStyle,
            string Weight,
            string Color);
    }
}
