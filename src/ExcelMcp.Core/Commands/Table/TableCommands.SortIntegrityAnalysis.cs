using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.Table;

public partial class TableCommands
{
    private static void InspectAdjacentTableDataAndFormulas(
        Excel.Range tableRange,
        Excel.Range? dataBodyRange,
        List<TablePreflightFinding> findings,
        CancellationToken cancellationToken)
    {
        Excel.Worksheet? worksheet = null;
        Excel.Range? tableRows = null;
        Excel.Range? tableColumns = null;
        Excel.Range? dataRows = null;
        Excel.Range? usedRange = null;
        Excel.Range? usedColumns = null;
        try
        {
            worksheet = tableRange.Worksheet;
            tableRows = tableRange.Rows;
            tableColumns = tableRange.Columns;
            usedRange = worksheet.UsedRange;
            usedColumns = usedRange.Columns;

            int tableFirstRow = Convert.ToInt32(tableRange.Row, CultureInfo.InvariantCulture);
            int tableLastRow = tableFirstRow
                + Convert.ToInt32(tableRows.Count, CultureInfo.InvariantCulture) - 1;
            int tableFirstColumn = Convert.ToInt32(tableRange.Column, CultureInfo.InvariantCulture);
            int tableLastColumn = tableFirstColumn
                + Convert.ToInt32(tableColumns.Count, CultureInfo.InvariantCulture) - 1;
            int usedFirstColumn = Convert.ToInt32(usedRange.Column, CultureInfo.InvariantCulture);
            int usedLastColumn = usedFirstColumn
                + Convert.ToInt32(usedColumns.Count, CultureInfo.InvariantCulture) - 1;
            int dataFirstRow = 0;
            int dataLastRow = -1;
            if (dataBodyRange is not null)
            {
                dataRows = dataBodyRange.Rows;
                dataFirstRow = Convert.ToInt32(dataBodyRange.Row, CultureInfo.InvariantCulture);
                dataLastRow = dataFirstRow
                    + Convert.ToInt32(dataRows.Count, CultureInfo.InvariantCulture) - 1;
            }

            var formulaAddresses = new List<string>();
            List<int> leftColumns = ScanAdjacentColumns(
                worksheet,
                tableFirstColumn - 1,
                usedFirstColumn,
                step: -1,
                tableFirstRow,
                tableLastRow,
                cancellationToken);
            List<int> rightColumns = ScanAdjacentColumns(
                worksheet,
                tableLastColumn + 1,
                usedLastColumn,
                step: 1,
                tableFirstRow,
                tableLastRow,
                cancellationToken);
            CollectExternalFormulaAddresses(
                worksheet,
                usedFirstColumn,
                usedLastColumn,
                tableFirstColumn,
                tableLastColumn,
                dataFirstRow,
                dataLastRow,
                formulaAddresses,
                cancellationToken);

            var excludedRanges = new List<string>();
            if (leftColumns.Count > 0)
            {
                excludedRanges.Add(GetAbsoluteRangeAddress(
                    leftColumns.Min(),
                    tableFirstRow,
                    leftColumns.Max(),
                    tableLastRow));
            }

            if (rightColumns.Count > 0)
            {
                excludedRanges.Add(GetAbsoluteRangeAddress(
                    rightColumns.Min(),
                    tableFirstRow,
                    rightColumns.Max(),
                    tableLastRow));
            }

            if (excludedRanges.Count > 0)
            {
                findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.ExcludedContiguousColumns,
                    Severity = TablePreflightSeverity.Warning,
                    IsHeuristic = true,
                    Addresses = excludedRanges,
                    Message = "Populated worksheet columns directly beside the table will not move with its rows.",
                    Remediation = "Confirm that these columns are separate data, or resize the table to include them before sorting."
                });
            }

            if (formulaAddresses.Count > 0)
            {
                findings.Add(new TablePreflightFinding
                {
                    Kind = TablePreflightFindingKind.RowAssociatedFormulaOutsideTable,
                    Severity = TablePreflightSeverity.Warning,
                    IsHeuristic = true,
                    Addresses = formulaAddresses,
                    Message = "These formulas are outside the table but align with its data rows and may be row-associated.",
                    Remediation = "Confirm that these formulas may remain fixed while table rows move, or include their data in the table."
                });
            }
        }
        finally
        {
            ComUtilities.Release(ref usedColumns);
            ComUtilities.Release(ref usedRange);
            ComUtilities.Release(ref dataRows);
            ComUtilities.Release(ref tableColumns);
            ComUtilities.Release(ref tableRows);
            ComUtilities.Release(ref worksheet);
        }
    }

    private static List<int> ScanAdjacentColumns(
        Excel.Worksheet worksheet,
        int startColumn,
        int boundaryColumn,
        int step,
        int tableFirstRow,
        int tableLastRow,
        CancellationToken cancellationToken)
    {
        var populatedColumns = new List<int>();
        for (int column = startColumn;
             step < 0 ? column >= boundaryColumn : column <= boundaryColumn;
             column += step)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Excel.Range? columnRange = null;
            try
            {
                columnRange = worksheet.Range[
                    GetAbsoluteRangeAddress(column, tableFirstRow, column, tableLastRow)];
                object formulas = columnRange.FormulaR1C1;
                bool populated = false;
                for (int rowOffset = 0; rowOffset <= tableLastRow - tableFirstRow; rowOffset++)
                {
                    string content = Convert.ToString(
                        GetMatrixValue(formulas, rowOffset, 0),
                        CultureInfo.InvariantCulture) ?? string.Empty;
                    if (content.Length == 0)
                    {
                        continue;
                    }

                    populated = true;
                }

                if (!populated)
                {
                    break;
                }

                populatedColumns.Add(column);
            }
            finally
            {
                ComUtilities.Release(ref columnRange);
            }
        }

        return populatedColumns;
    }

    private static void CollectExternalFormulaAddresses(
        Excel.Worksheet worksheet,
        int usedFirstColumn,
        int usedLastColumn,
        int tableFirstColumn,
        int tableLastColumn,
        int dataFirstRow,
        int dataLastRow,
        List<string> formulaAddresses,
        CancellationToken cancellationToken)
    {
        if (dataLastRow < dataFirstRow)
        {
            return;
        }

        Excel.Range? formulaRange = null;
        try
        {
            formulaRange = worksheet.Range[
                GetAbsoluteRangeAddress(usedFirstColumn, dataFirstRow, usedLastColumn, dataLastRow)];
            object formulas = formulaRange.FormulaR1C1;
            int rowCount = dataLastRow - dataFirstRow + 1;
            int columnCount = usedLastColumn - usedFirstColumn + 1;
            for (int rowOffset = 0; rowOffset < rowCount; rowOffset++)
            {
                for (int columnOffset = 0; columnOffset < columnCount; columnOffset++)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    int worksheetColumn = usedFirstColumn + columnOffset;
                    if (worksheetColumn >= tableFirstColumn && worksheetColumn <= tableLastColumn)
                    {
                        continue;
                    }

                    string content = Convert.ToString(
                        GetMatrixValue(formulas, rowOffset, columnOffset),
                        CultureInfo.InvariantCulture) ?? string.Empty;
                    if (content.StartsWith('='))
                    {
                        formulaAddresses.Add(GetAbsoluteAddress(
                            worksheetColumn,
                            dataFirstRow + rowOffset));
                    }
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref formulaRange);
        }
    }
}
