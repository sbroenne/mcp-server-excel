using System.Globalization;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Commands.Range;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

/// <summary>
/// Table creation safety checks.
/// </summary>
public partial class TableCommands
{
    private const long MaxSortSensitiveFormulaScanCells = 100_000;

    private static readonly Regex A1ReferenceRegex = new(
        @"(?<![A-Z0-9_])(?<column>\$?[A-Z]{1,3})(?<row>\$?\d+)(?![A-Z0-9_])",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    /// <inheritdoc />
    public TablePreflightResult Preflight(
        IExcelBatch batch,
        string sheetName,
        string tableName,
        string rangeAddress,
        bool hasHeaders = true)
    {
        ValidateCreateInputs(sheetName, tableName, rangeAddress);

        return batch.Execute((ctx, ct) =>
        {
            Excel.Worksheet? sheet = null;
            Excel.Range? effectiveRange = null;
            try
            {
                sheet = ComUtilities.FindSheet(ctx.Book, sheetName)
                    ?? throw new InvalidOperationException($"Sheet '{sheetName}' not found.");
                effectiveRange = ResolveEffectiveRange(sheet, rangeAddress);

                return AnalyzePreflight(
                    ctx.Book,
                    effectiveRange,
                    batch.WorkbookPath,
                    sheetName,
                    tableName,
                    rangeAddress,
                    hasHeaders,
                    ct);
            }
            finally
            {
                ComUtilities.Release(ref effectiveRange);
                ComUtilities.Release(ref sheet);
            }
        });
    }

    private static void ValidateCreateInputs(string sheetName, string tableName, string rangeAddress)
    {
        ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
        ArgumentException.ThrowIfNullOrWhiteSpace(rangeAddress);
        ValidateTableName(tableName);
    }

    private static Excel.Range ResolveEffectiveRange(Excel.Worksheet sheet, string rangeAddress)
    {
        Excel.Range? requestedRange = null;
        Excel.Range? currentRegion = null;
        try
        {
            requestedRange = sheet.Range[rangeAddress];
            if (Convert.ToInt64(requestedRange.CountLarge, CultureInfo.InvariantCulture) == 1)
            {
                currentRegion = requestedRange.CurrentRegion;
                if (Convert.ToInt64(currentRegion.CountLarge, CultureInfo.InvariantCulture) > 1)
                {
                    ComUtilities.Release(ref requestedRange);
                    requestedRange = currentRegion;
                    currentRegion = null;
                }
            }

            Excel.Range result = requestedRange;
            requestedRange = null;
            return result;
        }
        finally
        {
            ComUtilities.Release(ref currentRegion);
            ComUtilities.Release(ref requestedRange);
        }
    }

    private static TablePreflightResult AnalyzePreflight(
        Excel.Workbook workbook,
        Excel.Range effectiveRange,
        string workbookPath,
        string sheetName,
        string tableName,
        string requestedRange,
        bool hasHeaders,
        CancellationToken cancellationToken)
    {
        var result = new TablePreflightResult
        {
            Success = true,
            FilePath = workbookPath,
            SheetName = sheetName,
            TableName = tableName,
            RequestedRange = requestedRange,
            EffectiveRange = effectiveRange.Address
        };

        if (TableExists(workbook, tableName))
        {
            result.Findings.Add(new TablePreflightFinding
            {
                Kind = TablePreflightFindingKind.TableNameExists,
                Severity = TablePreflightSeverity.Blocker,
                Message = $"A table named '{tableName}' already exists in this workbook.",
                Remediation = "Choose a unique table name or use the existing table."
            });
        }

        bool? isMerged = RangeMergeDiscovery.GetMergeCellsState(effectiveRange.MergeCells);
        List<string> mergedRanges = isMerged == false
            ? []
            : RangeMergeDiscovery.CollectMergedRanges(effectiveRange, isMerged, cancellationToken);
        if (mergedRanges.Count > 0)
        {
            result.Findings.Add(new TablePreflightFinding
            {
                Kind = TablePreflightFindingKind.MergedCells,
                Severity = TablePreflightSeverity.Blocker,
                Addresses = [.. mergedRanges],
                Message = "The proposed table range intersects merged cells, which Excel tables cannot safely preserve.",
                Remediation = "Unmerge these cells and repeat any shared labels in each resulting cell before creating the table."
            });
        }

        InspectHeaders(effectiveRange, hasHeaders, result);
        InspectExcludedContiguousColumns(effectiveRange, result);
        InspectSortSensitiveFormulas(effectiveRange, hasHeaders, result, cancellationToken);

        result.SafeToCreate = result.Findings.All(
            finding => finding.Severity != TablePreflightSeverity.Blocker);
        return result;
    }

    private static void InspectHeaders(
        Excel.Range effectiveRange,
        bool hasHeaders,
        TablePreflightResult result)
    {
        if (!hasHeaders)
        {
            return;
        }

        Excel.Range? rows = null;
        Excel.Range? columns = null;
        Excel.Range? headerRange = null;
        try
        {
            rows = effectiveRange.Rows;
            columns = effectiveRange.Columns;
            headerRange = rows[1];

            object values = headerRange.Value2;
            int columnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
            int headerRow = Convert.ToInt32(effectiveRange.Row, CultureInfo.InvariantCulture);
            int firstColumn = Convert.ToInt32(effectiveRange.Column, CultureInfo.InvariantCulture);
            var blankAddresses = new List<string>();
            var addressesByHeader = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

            for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
            {
                string address = GetAbsoluteAddress(firstColumn + columnOffset, headerRow);
                string header = Convert.ToString(
                    GetMatrixValue(values, 0, columnOffset),
                    CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;

                if (header.Length == 0)
                {
                    blankAddresses.Add(address);
                    continue;
                }

                if (!addressesByHeader.TryGetValue(header, out List<string>? addresses))
                {
                    addresses = [];
                    addressesByHeader.Add(header, addresses);
                }

                addresses.Add(address);
            }

            if (blankAddresses.Count > 0)
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.BlankHeaders,
                    Severity = TablePreflightSeverity.Blocker,
                    Addresses = blankAddresses,
                    Message = "One or more table headers are blank, so Excel would generate names silently.",
                    Remediation = "Enter a unique, descriptive header in each listed cell before creating the table."
                });
            }

            List<string> duplicateAddresses = addressesByHeader
                .Where(group => group.Value.Count > 1)
                .SelectMany(group => group.Value)
                .ToList();
            if (duplicateAddresses.Count > 0)
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.DuplicateHeaders,
                    Severity = TablePreflightSeverity.Blocker,
                    Addresses = duplicateAddresses,
                    Message = "Two or more table headers have the same name after trimming spaces and ignoring case.",
                    Remediation = "Rename the listed headers so every table column has a unique name."
                });
            }
        }
        finally
        {
            ComUtilities.Release(ref headerRange);
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
        }
    }

    private static void InspectExcludedContiguousColumns(
        Excel.Range effectiveRange,
        TablePreflightResult result)
    {
        Excel.Range? effectiveColumns = null;
        Excel.Range? currentRegion = null;
        Excel.Range? regionRows = null;
        Excel.Range? regionColumns = null;
        try
        {
            effectiveColumns = effectiveRange.Columns;
            currentRegion = effectiveRange.CurrentRegion;
            regionRows = currentRegion.Rows;
            regionColumns = currentRegion.Columns;

            int effectiveFirstColumn = Convert.ToInt32(effectiveRange.Column, CultureInfo.InvariantCulture);
            int effectiveLastColumn = effectiveFirstColumn
                + Convert.ToInt32(effectiveColumns.Count, CultureInfo.InvariantCulture) - 1;
            int regionFirstColumn = Convert.ToInt32(currentRegion.Column, CultureInfo.InvariantCulture);
            int regionLastColumn = regionFirstColumn
                + Convert.ToInt32(regionColumns.Count, CultureInfo.InvariantCulture) - 1;
            int regionFirstRow = Convert.ToInt32(currentRegion.Row, CultureInfo.InvariantCulture);
            int regionLastRow = regionFirstRow
                + Convert.ToInt32(regionRows.Count, CultureInfo.InvariantCulture) - 1;
            var excludedRanges = new List<string>();

            if (regionFirstColumn < effectiveFirstColumn)
            {
                excludedRanges.Add(GetAbsoluteRangeAddress(
                    regionFirstColumn,
                    regionFirstRow,
                    effectiveFirstColumn - 1,
                    regionLastRow));
            }

            if (regionLastColumn > effectiveLastColumn)
            {
                excludedRanges.Add(GetAbsoluteRangeAddress(
                    effectiveLastColumn + 1,
                    regionFirstRow,
                    regionLastColumn,
                    regionLastRow));
            }

            if (excludedRanges.Count > 0)
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.ExcludedContiguousColumns,
                    Severity = TablePreflightSeverity.Warning,
                    IsHeuristic = true,
                    Addresses = excludedRanges,
                    Message = "The same contiguous data region contains populated columns outside the proposed table range.",
                    Remediation = "Expand the table range to include these columns, or confirm that they are separate data."
                });
            }
        }
        finally
        {
            ComUtilities.Release(ref regionColumns);
            ComUtilities.Release(ref regionRows);
            ComUtilities.Release(ref currentRegion);
            ComUtilities.Release(ref effectiveColumns);
        }
    }

    private static void InspectSortSensitiveFormulas(
        Excel.Range effectiveRange,
        bool hasHeaders,
        TablePreflightResult result,
        CancellationToken cancellationToken)
    {
        Excel.Range? rows = null;
        Excel.Range? columns = null;
        try
        {
            rows = effectiveRange.Rows;
            columns = effectiveRange.Columns;
            int rowCount = Convert.ToInt32(rows.Count, CultureInfo.InvariantCulture);
            int columnCount = Convert.ToInt32(columns.Count, CultureInfo.InvariantCulture);
            int firstDataRowOffset = hasHeaders ? 1 : 0;
            if (rowCount <= firstDataRowOffset)
            {
                return;
            }

            long cellCount = Convert.ToInt64(effectiveRange.CountLarge, CultureInfo.InvariantCulture);
            if (cellCount > MaxSortSensitiveFormulaScanCells)
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.FormulaScanSkipped,
                    Severity = TablePreflightSeverity.Warning,
                    IsHeuristic = true,
                    Message = $"Formula sorting risk analysis was skipped because the proposed range contains " +
                        $"{cellCount.ToString("N0", CultureInfo.InvariantCulture)} cells, exceeding the bounded scan limit " +
                        $"of {MaxSortSensitiveFormulaScanCells.ToString("N0", CultureInfo.InvariantCulture)} cells.",
                    Remediation = "Run preflight on a smaller range or review the table formulas manually before sorting."
                });
                return;
            }

            int firstRow = Convert.ToInt32(effectiveRange.Row, CultureInfo.InvariantCulture);
            int firstColumn = Convert.ToInt32(effectiveRange.Column, CultureInfo.InvariantCulture);
            int lastColumn = firstColumn + columnCount - 1;
            object formulas = effectiveRange.Formula;
            var warningAddresses = new List<string>();

            for (int rowOffset = firstDataRowOffset; rowOffset < rowCount; rowOffset++)
            {
                for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    string formula = Convert.ToString(
                        GetMatrixValue(formulas, rowOffset, columnOffset),
                        CultureInfo.InvariantCulture) ?? string.Empty;
                    if (!formula.StartsWith('='))
                    {
                        continue;
                    }

                    int formulaRow = firstRow + rowOffset;
                    if (HasSortSensitiveReference(formula, formulaRow, firstColumn, lastColumn))
                    {
                        warningAddresses.Add(GetAbsoluteAddress(firstColumn + columnOffset, formulaRow));
                    }
                }
            }

            if (warningAddresses.Count > 0)
            {
                result.Findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.SortSensitiveFormula,
                    Severity = TablePreflightSeverity.Warning,
                    IsHeuristic = true,
                    Addresses = warningAddresses,
                    Message = "These formulas use fixed-row or cross-row A1 references that may no longer align after sorting the table.",
                    Remediation = "Use structured references where possible, or confirm the formulas still point to the intended rows after sorting."
                });
            }
        }
        finally
        {
            ComUtilities.Release(ref columns);
            ComUtilities.Release(ref rows);
        }
    }

    internal static bool HasSortSensitiveReference(
        string formula,
        int formulaRow,
        int firstTableColumn,
        int lastTableColumn)
    {
        foreach (Match match in EnumerateUnquotedA1References(formula))
        {
            string rowToken = match.Groups["row"].Value;
            if (!int.TryParse(rowToken.TrimStart('$'), NumberStyles.None, CultureInfo.InvariantCulture, out int referencedRow))
            {
                continue;
            }

            string columnToken = match.Groups["column"].Value;
            int referencedColumn = GetColumnIndex(columnToken.TrimStart('$'));
            bool isSheetQualified = match.Index > 0 && formula[match.Index - 1] == '!';
            if (rowToken.StartsWith('$')
                || referencedRow != formulaRow
                || referencedColumn < firstTableColumn
                || referencedColumn > lastTableColumn
                || isSheetQualified)
            {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<Match> EnumerateUnquotedA1References(string formula)
    {
        int scanIndex = 0;
        bool isInStringLiteral = false;

        foreach (Match match in A1ReferenceRegex.Matches(formula))
        {
            while (scanIndex < match.Index)
            {
                if (formula[scanIndex] != '"')
                {
                    scanIndex++;
                    continue;
                }

                if (isInStringLiteral
                    && scanIndex + 1 < match.Index
                    && formula[scanIndex + 1] == '"')
                {
                    scanIndex += 2;
                    continue;
                }

                isInStringLiteral = !isInStringLiteral;
                scanIndex++;
            }

            scanIndex = match.Index + match.Length;
            if (!isInStringLiteral)
            {
                yield return match;
            }
        }
    }

    private static int GetColumnIndex(string columnName)
    {
        int column = 0;
        foreach (char character in columnName)
        {
            column = checked(column * 26 + char.ToUpperInvariant(character) - 'A' + 1);
        }

        return column;
    }

    private static object? GetMatrixValue(object valueOrArray, int rowOffset, int columnOffset)
    {
        if (valueOrArray is not Array values || values.Rank != 2)
        {
            return rowOffset == 0 && columnOffset == 0 ? valueOrArray : null;
        }

        return values.GetValue(
            values.GetLowerBound(0) + rowOffset,
            values.GetLowerBound(1) + columnOffset);
    }

    private static string GetAbsoluteRangeAddress(
        int firstColumn,
        int firstRow,
        int lastColumn,
        int lastRow) =>
        $"{GetAbsoluteAddress(firstColumn, firstRow)}:{GetAbsoluteAddress(lastColumn, lastRow)}";

    private static string GetAbsoluteAddress(int column, int row) =>
        $"${GetColumnName(column)}${row}";

    private static string GetColumnName(int column)
    {
        string name = string.Empty;
        while (column > 0)
        {
            column--;
            name = Convert.ToChar('A' + column % 26) + name;
            column /= 26;
        }

        return name;
    }
}
