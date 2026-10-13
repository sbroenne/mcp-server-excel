using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Shared helper methods for PivotTable operations.
/// Centralizes common patterns to avoid cargo cult duplication.
/// </summary>
internal static class PivotTableHelpers
{
    internal static bool IsAllOlapReportFilterMember(
        Excel.CubeField cubeField, string memberName, CancellationToken cancellationToken)
    {
        const string adoUnavailableMessage =
            "Excel did not expose an active ADO session to identify the OLAP All member. " +
            "If a refresh is authorized, use pivottable(refresh) for this PivotTable " +
            "(CLI: excelcli pivottable refresh), then retry the read.";
        Excel.PivotTable? pivot = null;
        Excel.PivotCache? cache = null;
        Excel.WorkbookConnection? workbookConnection = null;
        dynamic? adoConnection = null;
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            pivot = (Excel.PivotTable)cubeField.Parent;
            cache = pivot.PivotCache();
            string? cubeName = cache.CommandText switch
            {
                string text => text,
                string[] { Length: 1 } texts => texts[0],
                object[] { Length: 1 } values => Convert.ToString(values[0], CultureInfo.InvariantCulture),
                _ => null
            };
            if (string.IsNullOrWhiteSpace(cubeName))
                throw new NotSupportedException("Excel did not expose the selected OLAP cube name.");

            // ADO schema rowsets have no referenced PIA. AllItemsVisible does not
            // distinguish All from a selected member in single-selection mode.
            try
            {
                adoConnection = cache.ADOConnection;
            }
            catch (System.Runtime.InteropServices.COMException ex) when (CreatedObjectFailure.CanReport(ex))
            {
                throw new NotSupportedException(adoUnavailableMessage, ex);
            }
            if (adoConnection is null)
                throw new NotSupportedException(adoUnavailableMessage);

            workbookConnection = cache.WorkbookConnection;
            if (workbookConnection is null)
                throw new NotSupportedException("Excel did not expose the workbook connection for this OLAP PivotTable.");

            List<IReadOnlyDictionary<string, object?>> rows;
            try
            {
                rows = ConnectionCommands.OpenSchemaRowset(
                    adoConnection,
                    ExternalOlapSchemaMapper.BuildHierarchyRequest(cubeName, cubeField.Name),
                    workbookConnection.Name,
                    cancellationToken,
                    maximumRows: 2);
            }
            catch (InvalidOperationException ex) when (CreatedObjectFailure.CanReport(ex))
            {
                throw new NotSupportedException(ex.Message, ex);
            }
            if (rows.Count != 1 || !rows[0].TryGetValue("ALL_MEMBER", out object? allMember))
                throw new NotSupportedException("Excel did not expose the All-member metadata for this OLAP hierarchy.");

            return allMember is not null
                && string.Equals(memberName, Convert.ToString(allMember, CultureInfo.InvariantCulture),
                    StringComparison.Ordinal);
        }
        finally
        {
            ComUtilities.Release(ref workbookConnection);
            ComUtilities.Release(ref adoConnection);
            ComUtilities.Release(ref cache);
            ComUtilities.Release(ref pivot);
        }
    }

    /// <summary>
    /// Detects Excel's silent collapse of a blocked PivotTable to a single #SPILL! cell.
    /// Empty or otherwise small PivotTables are not layout failures.
    /// </summary>
    internal static string? GetLayoutError(Excel.PivotTable pivot)
    {
        Excel.Range? range = null;
        try
        {
            range = pivot.TableRange1;
            if (range is not null && Convert.ToDouble((object)range.CountLarge, CultureInfo.InvariantCulture) == 1
                && ExcelErrorMapper.TryGet((object?)range.Value2, out _, out var error)
                && error.Name == "#SPILL!")
            {
                return $"PivotTable '{pivot.Name}' is not drawn: its range {range.Address} contains #SPILL!. " +
                    "The layout may overlap another PivotTable or be blocked by occupied cells.";
            }
            return null;
        }
        finally
        {
            ComUtilities.Release(ref range);
        }
    }

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
#pragma warning disable CA1031 // COM interop: CubeFields property doesn't exist on non-OLAP PivotTables
        catch (System.Runtime.InteropServices.COMException ex) when (CreatedObjectFailure.CanReport(ex))
        {
            // CubeFields property not available or failed - not an OLAP PivotTable
            return false;
        }
#pragma warning restore CA1031
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
#pragma warning disable CA1031 // COM interop: CubeFields property doesn't exist on non-OLAP PivotTables
        catch (System.Runtime.InteropServices.COMException)
        {
            // CubeFields property not available - not an OLAP PivotTable
            // cubeFields already null from initialization
            return false;
        }
#pragma warning restore CA1031
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

    /// <summary>
    /// Detects the data type of a PivotField by sampling its values.
    /// </summary>
    /// <param name="field">The PivotField COM object</param>
    /// <returns>Data type string: "Date", "Number", "Boolean", "Text", or "Unknown"</returns>
    public static string DetectFieldDataType(dynamic field)
    {
        string? fieldDataType = TryDetectFromPivotFieldDataType(field);
        if (fieldDataType != null)
        {
            return fieldDataType;
        }

        dynamic? pivotItems = null;
        try
        {
            pivotItems = field.PivotItems;
            var sampleValues = new List<object?>();

            int sampleCount = Math.Min(10, pivotItems.Count);
            for (int i = 1; i <= sampleCount; i++)
            {
                dynamic? item = null;
                try
                {
                    item = pivotItems.Item(i);
                    var value = item.Value;
                    if (value != null)
                        sampleValues.Add(value);
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }

            // Analyze sample values
            if (sampleValues.Count == 0)
                return "Unknown";

            if (sampleValues.All(v => DateTime.TryParse(v?.ToString(), out _)))
                return "Date";
            if (sampleValues.All(v => double.TryParse(v?.ToString(), out _)))
                return "Number";
            if (sampleValues.All(v => bool.TryParse(v?.ToString(), out _)))
                return "Boolean";

            return "Text";
        }
#pragma warning disable CA1031 // COM interop: PivotItems may not be accessible on all field types
        catch (System.Runtime.InteropServices.COMException)
        {
            // PivotItems access failed - cannot determine data type
            return "Unknown";
        }
#pragma warning restore CA1031
        finally
        {
            ComUtilities.Release(ref pivotItems);
        }
    }

    private static string? TryDetectFromPivotFieldDataType(dynamic field)
    {
        try
        {
            // Excel reports calculated fields as text even though their results are numeric.
            if (((Microsoft.Office.Interop.Excel.PivotField)field).IsCalculated)
                return "Number";

            int dataType = Convert.ToInt32(field.DataType);
            return dataType switch
            {
                XlPivotFieldDataType.xlDate => "Date",
                XlPivotFieldDataType.xlNumber => "Number",
                XlPivotFieldDataType.xlText => "Text",
                _ => null
            };
        }
        catch (System.Runtime.InteropServices.COMException)
        {
            return null;
        }
    }

    /// <summary>
    /// Gets unique values from a PivotField for filtering purposes.
    /// Iterates through PivotItems to collect available values.
    /// </summary>
    /// <param name="field">The PivotField COM object</param>
    /// <returns>List of unique value strings from the field</returns>
    public static List<string> GetFieldUniqueValues(dynamic field)
    {
        var values = new List<string>();
        dynamic? pivotItems = null;
        try
        {
            pivotItems = field.PivotItems;
            for (int i = 1; i <= pivotItems.Count; i++)
            {
                dynamic? item = null;
                try
                {
                    item = pivotItems.Item(i);
                    string itemName = item.Name?.ToString() ?? string.Empty;
                    if (!string.IsNullOrEmpty(itemName))
                        values.Add(itemName);
                }
                finally
                {
                    ComUtilities.Release(ref item);
                }
            }
        }
#pragma warning disable CA1031 // COM interop: PivotItems may not be accessible on all field types
        catch (System.Runtime.InteropServices.COMException)
        {
            // PivotItems access failed - return partial list
        }
#pragma warning restore CA1031
        finally
        {
            ComUtilities.Release(ref pivotItems);
        }
        return values;
    }
}
