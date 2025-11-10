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
    public static IPivotTableFieldStrategy GetStrategy(dynamic pivot)
    {
        if (pivot == null)
            throw new InvalidOperationException("PivotTable object is null");

        // Try OLAP first (more specific)
        try
        {
            dynamic? cubeFields = pivot.CubeFields;
            if (cubeFields != null && cubeFields.Count > 0)
            {
                ComUtilities.Release(ref cubeFields);
                return _strategies[0]; // OlapPivotTableFieldStrategy
            }
            ComUtilities.Release(ref cubeFields);
        }
        catch
        {
            // Not OLAP
        }

        // Fall back to Regular
        try
        {
            dynamic? pivotFields = pivot.PivotFields;
            if (pivotFields != null)
            {
                ComUtilities.Release(ref pivotFields);
                return _strategies[1]; // RegularPivotTableFieldStrategy
            }
            ComUtilities.Release(ref pivotFields);
        }
        catch
        {
            // Not Regular either
        }

        throw new InvalidOperationException("No strategy found for PivotTable type. Unable to determine if OLAP or Regular PivotTable.");
    }
}
