using Excel = Microsoft.Office.Interop.Excel;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// PivotTable field management operations
/// </summary>
public partial class PivotTableCommands
{
    /// <summary>
    /// Lists all fields in a PivotTable
    /// </summary>
    public PivotFieldListResult ListFields(IExcelBatch batch, string pivotTableName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;
            Excel.PivotFields? pivotFields = null;
            Excel.CubeFields? cubeFields = null;

            try
            {
                pivot = FindPivotTable(ctx.Book, pivotTableName);
                bool isOlap = PivotTableHelpers.TryGetCubeFields(pivot, out cubeFields);
                // For OLAP PivotTables, use CubeFields instead of PivotFields
                if (isOlap && cubeFields is not null)
                {
                    var result = ListCubeFieldsAsync(cubeFields, batch.WorkbookPath);
                    result.ValueFields = ReadAllValueFieldCalculations(
                        (Microsoft.Office.Interop.Excel.PivotTable)pivot, batch.WorkbookPath, ct);
                    return result;
                }
                else
                {
                    // Regular PivotTable - use PivotFields
                    pivotFields = (Excel.PivotFields)pivot.PivotFields();
                    var result = ListRegularFieldsAsync(pivotFields, batch.WorkbookPath);
                    result.ValueFields = ReadAllValueFieldCalculations(
                        (Microsoft.Office.Interop.Excel.PivotTable)pivot, batch.WorkbookPath, ct);
                    return result;
                }
            }
            finally
            {
                ComUtilities.Release(ref cubeFields);
                ComUtilities.Release(ref pivotFields);
                ComUtilities.Release(ref pivot);
            }
        });
    }

    /// <summary>
    /// Lists fields from OLAP/Data Model PivotTable using CubeFields
    /// </summary>
    private static PivotFieldListResult ListCubeFieldsAsync(Excel.CubeFields cubeFields, string filePath)
    {
        return new PivotFieldListResult
        {
            Success = true,
            Fields = GetCubeFieldsInfo(cubeFields),
            FilePath = filePath
        };
    }

    /// <summary>
    /// Lists fields from regular PivotTable using PivotFields
    /// </summary>
    private static PivotFieldListResult ListRegularFieldsAsync(Excel.PivotFields pivotFields, string filePath)
    {
        var fields = new List<PivotFieldInfo>();

        int fieldCount = pivotFields.Count;

        for (int i = 1; i <= fieldCount; i++)
        {
            Excel.PivotField? field = null;
            try
            {
                field = pivotFields.Item(i);

                // Get field name with defensive handling (can throw on Data Model fields)
                string fieldName;
                try
                {
                    object? sourceName = field.SourceName;
                    if (PivotTableHelpers.IsValuesPlaceholderField(field, sourceName))
                        continue; // The synthetic Values field exposes an error VARIANT, not a source name.
                    fieldName = sourceName?.ToString() ?? field.Name?.ToString() ?? $"Field{i}";
                }
                catch (System.Runtime.InteropServices.COMException ex) when (ComErrors.IsMemberUnavailable(ex))
                {
                    fieldName = $"Field{i}";
                }

                // Get orientation with defensive handling
                int orientation;
                try
                {
                    orientation = (int)field.Orientation;
                }
                catch (System.Runtime.InteropServices.COMException ex) when (ComErrors.IsMemberUnavailable(ex))
                {
                    orientation = XlPivotFieldOrientation.xlHidden;
                }

                var fieldInfo = new PivotFieldInfo
                {
                    Name = fieldName,
                    Area = orientation switch
                    {
                        XlPivotFieldOrientation.xlRowField => PivotFieldArea.Row,
                        XlPivotFieldOrientation.xlColumnField => PivotFieldArea.Column,
                        XlPivotFieldOrientation.xlPageField => PivotFieldArea.Filter,
                        XlPivotFieldOrientation.xlDataField => PivotFieldArea.Value,
                        _ => PivotFieldArea.Hidden
                    }
                };

                // CustomName - defensive
                try
                {
                    fieldInfo.CustomName = field.Caption?.ToString() ?? string.Empty;
                }
                catch (System.Runtime.InteropServices.COMException ex) when (ComErrors.IsMemberUnavailable(ex))
                {
                    fieldInfo.CustomName = string.Empty;
                }

                // Position - defensive
                try
                {
                    fieldInfo.Position = orientation != XlPivotFieldOrientation.xlHidden ? Convert.ToInt32((object)field.Position, System.Globalization.CultureInfo.InvariantCulture) : 0;
                }
                catch (System.Runtime.InteropServices.COMException ex) when (ComErrors.IsMemberUnavailable(ex))
                {
                    fieldInfo.Position = 0;
                }

                // DataType - defensive
                try
                {
                    fieldInfo.DataType = PivotTableHelpers.DetectFieldDataType(field);
                }
                catch (System.Runtime.InteropServices.COMException ex) when (ComErrors.IsMemberUnavailable(ex))
                {
                    fieldInfo.DataType = "Unknown";
                }

                // Get function for value fields - defensive
                if (orientation == XlPivotFieldOrientation.xlDataField)
                {
                    try
                    {
                        int comFunction = (int)field.Function;
                        fieldInfo.Function = PivotTableHelpers.GetAggregationFunctionFromCom(comFunction);
                    }
                    catch (System.Runtime.InteropServices.COMException ex) when (ComErrors.IsMemberUnavailable(ex))
                    {
                        fieldInfo.Function = AggregationFunction.Sum; // Default
                    }
                }

                fields.Add(fieldInfo);
            }
            finally
            {
                ComUtilities.Release(ref field);
            }
        }

        return new PivotFieldListResult
        {
            Success = true,
            Fields = fields,
            FilePath = filePath
        };
    }

    /// <summary>
    /// Adds a field to the Row area
    /// </summary>
    public PivotFieldResult AddRowField(IExcelBatch batch, string pivotTableName,
        string fieldName, int? position = null)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.AddRowField(pivot, fieldName, position, batch.WorkbookPath));

    /// <summary>
    /// Adds a field to the Column area
    /// </summary>
    public PivotFieldResult AddColumnField(IExcelBatch batch, string pivotTableName,
        string fieldName, int? position = null)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.AddColumnField(pivot, fieldName, position, batch.WorkbookPath));

    /// <summary>
    /// Adds a field to the Values area with aggregation
    /// </summary>
    public PivotFieldResult AddValueField(IExcelBatch batch, string pivotTableName,
        string fieldName, AggregationFunction aggregationFunction = AggregationFunction.Sum,
        string? customName = null)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) =>
            {
                var result = strategy.AddValueField(pivot, fieldName, aggregationFunction, customName, batch.WorkbookPath);
                if (!result.Success)
                {
                    return result;
                }
                var layoutError = PivotTableHelpers.GetLayoutError((Microsoft.Office.Interop.Excel.PivotTable)pivot);
                if (layoutError is not null)
                {
                    throw new InvalidOperationException(
                        $"Value field '{fieldName}' was added, but {layoutError} " +
                        "The field change has not been rolled back. Inspect the PivotTable and nearby cells before deciding how to recover.");
                }
                return result;
            });

    /// <summary>
    /// Adds a field to the Filter area
    /// </summary>
    public PivotFieldResult AddFilterField(IExcelBatch batch, string pivotTableName,
        string fieldName)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.AddFilterField(pivot, fieldName, batch.WorkbookPath));

    /// <summary>
    /// Removes a field from any area
    /// </summary>
    public PivotFieldResult RemoveField(IExcelBatch batch, string pivotTableName,
        string fieldName)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.RemoveField(pivot, fieldName, batch.WorkbookPath));

    /// <summary>
    /// Sets the aggregation function for a value field
    /// </summary>
    public PivotFieldResult SetFieldFunction(IExcelBatch batch, string pivotTableName,
        string fieldName, AggregationFunction aggregationFunction)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.SetFieldFunction(pivot, fieldName, aggregationFunction, batch.WorkbookPath));

    /// <summary>
    /// Sets custom name for a field
    /// </summary>
    public PivotFieldResult SetFieldName(IExcelBatch batch, string pivotTableName,
        string fieldName, string customName)
        => ExecuteWithStrategy<PivotFieldResult>(batch, pivotTableName,
            (strategy, pivot) => strategy.SetFieldName(pivot, fieldName, customName, batch.WorkbookPath));

    /// <summary>
    /// Sets number format for a value field
    /// </summary>
    public PivotFieldResult SetFieldFormat(IExcelBatch batch, string pivotTableName,
        string fieldName, string numberFormat)
    {
        var invariantFormat = NumberFormatLiterals.PreserveCurrencyLiterals(numberFormat);
        return batch.Execute((ctx, ct) =>
        {
            Excel.PivotTable? pivot = null;

            pivot = FindPivotTable(ctx.Book, pivotTableName);

            try
            {
                // Use Strategy Pattern to delegate to appropriate implementation
                var strategy = PivotTableFieldStrategyFactory.GetStrategy(pivot);
                return strategy.SetFieldFormat(pivot, fieldName, invariantFormat, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref pivot);
            }
        });
    }
}
