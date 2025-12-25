using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Shared helper methods for PivotTable operations.
/// Centralizes common patterns to avoid cargo cult duplication.
/// </summary>
internal static class PivotTableHelpers
{
    /// <summary>
    /// Gets the area name for display purposes from a pivot field orientation.
    /// </summary>
    public static string GetAreaName(dynamic orientation)
    {
        int orientationValue = Convert.ToInt32(orientation);
        return orientationValue switch
        {
            XlPivotFieldOrientation.xlHidden => "Hidden",
            XlPivotFieldOrientation.xlRowField => "Row",
            XlPivotFieldOrientation.xlColumnField => "Column",
            XlPivotFieldOrientation.xlPageField => "Filter",
            XlPivotFieldOrientation.xlDataField => "Value",
            _ => $"Unknown({orientationValue})"
        };
    }

    /// <summary>
    /// Determines if a PivotTable is OLAP-based (Data Model/PowerPivot).
    /// OLAP PivotTables use CubeFields for field manipulation, while regular
    /// PivotTables use PivotFields.
    /// </summary>
    /// <param name="pivot">The PivotTable COM object</param>
    /// <returns>True if the PivotTable is OLAP-based, false otherwise</returns>
    /// <remarks>
    /// This helper consolidates the duplicated pattern:
    ///   cubeFields = pivot.CubeFields;
    ///   isOlap = cubeFields != null &amp;&amp; cubeFields.Count > 0;
    ///
    /// Note: Does NOT release cubeFields - caller may need them.
    /// Use TryGetCubeFields for patterns that need the COM object.
    /// </remarks>
    public static bool IsOlapPivotTable(dynamic pivot)
    {
        dynamic? cubeFields = null;
        try
        {
            cubeFields = pivot.CubeFields;
            return cubeFields != null && cubeFields.Count > 0;
        }
        catch
        {
            // CubeFields property not available or failed - not an OLAP PivotTable
            return false;
        }
        finally
        {
            ComUtilities.Release(ref cubeFields);
        }
    }

    /// <summary>
    /// Attempts to get CubeFields from a PivotTable for OLAP operations.
    /// Returns the cubeFields object along with the OLAP status.
    /// </summary>
    /// <param name="pivot">The PivotTable COM object</param>
    /// <param name="cubeFields">Output: The CubeFields collection (caller must release)</param>
    /// <returns>True if OLAP with valid CubeFields, false otherwise</returns>
    /// <remarks>
    /// IMPORTANT: Caller is responsible for releasing cubeFields via ComUtilities.Release().
    /// Use this when you need the cubeFields object for subsequent operations.
    /// </remarks>
    public static bool TryGetCubeFields(dynamic pivot, out dynamic? cubeFields)
    {
        cubeFields = null;
        try
        {
            cubeFields = pivot.CubeFields;
            return cubeFields != null && cubeFields.Count > 0;
        }
        catch
        {
            // CubeFields property not available - not an OLAP PivotTable
            // cubeFields already null from initialization
            return false;
        }
    }

    /// <summary>
    /// Converts Excel COM constant to AggregationFunction enum.
    /// </summary>
    public static AggregationFunction GetAggregationFunctionFromCom(int comFunction)
    {
        return comFunction switch
        {
            XlConsolidationFunction.xlSum => AggregationFunction.Sum,
            XlConsolidationFunction.xlCount => AggregationFunction.Count,
            XlConsolidationFunction.xlAverage => AggregationFunction.Average,
            XlConsolidationFunction.xlMax => AggregationFunction.Max,
            XlConsolidationFunction.xlMin => AggregationFunction.Min,
            XlConsolidationFunction.xlProduct => AggregationFunction.Product,
            XlConsolidationFunction.xlCountNums => AggregationFunction.CountNumbers,
            XlConsolidationFunction.xlStdDev => AggregationFunction.StdDev,
            XlConsolidationFunction.xlStdDevP => AggregationFunction.StdDevP,
            XlConsolidationFunction.xlVar => AggregationFunction.Var,
            XlConsolidationFunction.xlVarP => AggregationFunction.VarP,
            _ => throw new InvalidOperationException($"Unknown COM aggregation function: {comFunction}")
        };
    }
}
