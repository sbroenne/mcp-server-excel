using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Factory for creating appropriate PivotTable field strategy based on PivotTable type
/// </summary>
public static class PivotTableFieldStrategyFactory
{
    private static readonly List<IPivotTableFieldStrategy> _strategies = new()
    {
        new OlapPivotTableFieldStrategy(),      // Check OLAP first (more specific)
        new RegularPivotTableFieldStrategy()    // Fallback to regular
    };

    /// <summary>
    /// Gets the appropriate strategy for the given PivotTable
    /// </summary>
    /// <param name="pivot">The PivotTable object</param>
    /// <returns>Strategy that can handle this PivotTable type</returns>
    /// <exception cref="InvalidOperationException">If no strategy can handle the PivotTable</exception>
    public static IPivotTableFieldStrategy GetStrategy(Excel.PivotTable pivot)
    {
        if (pivot == null)
            throw new InvalidOperationException("PivotTable object is null");

        // Try OLAP first (more specific)
        if (PivotTableHelpers.IsOlapPivotTable(pivot))
        {
            return _strategies[0]; // OlapPivotTableFieldStrategy
        }

        // Fall back to Regular
        Excel.PivotFields? pivotFields = null;
        try
        {
            pivotFields = (Excel.PivotFields)pivot.PivotFields();
            if (pivotFields != null)
            {
                return _strategies[1]; // RegularPivotTableFieldStrategy
            }
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            // Not Regular either
        }
        finally
        {
            ComUtilities.Release(ref pivotFields);
        }

        throw new InvalidOperationException("No strategy found for PivotTable type. Unable to determine if OLAP or Regular PivotTable.");
    }
}

