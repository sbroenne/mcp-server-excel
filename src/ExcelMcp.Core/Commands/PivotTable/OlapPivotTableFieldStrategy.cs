using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Strategy for OLAP (Online Analytical Processing) PivotTable field operations.
/// Uses CubeFields API for Data Model-based PivotTables.
///
/// CRITICAL: In OLAP PivotTables, PivotFields do not exist until the corresponding
/// CubeField is added to the PivotTable. Must call CreatePivotFields() first.
/// Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.cubefield.createpivotfields
/// </summary>
public class OlapPivotTableFieldStrategy : IPivotTableFieldStrategy
{
    private static readonly char[] FieldNameSeparators = ['[', ']', '.'];
    /// <inheritdoc/>
    public bool CanHandle(dynamic pivot)
    {
        try
        {
            // OLAP/Data Model PivotTables have CubeFields collection
            // Note: Don't release COM objects here - PivotTable keeps them alive
            dynamic cubeFields = pivot.CubeFields;
            return cubeFields != null && cubeFields.Count > 0;
        }
        catch
        {
            return false;
        }
    }

    /// <inheritdoc/>
    public dynamic GetFieldForManipulation(dynamic pivot, string fieldName)
    {
        dynamic? cubeFields = null;
        dynamic? cubeField = null;
        try
        {
            cubeFields = pivot.CubeFields;

            // EXACT MATCH ONLY - no partial matching to avoid disambiguation bugs
            // The LLM knows exact field names, so partial matching only causes problems
            // (e.g., "ACR" incorrectly matching "[DisambiguationTable].[ACRTypeKey]")
            try
            {
                cubeField = cubeFields.Item(fieldName);
            }
            catch
            {
                // Field not found by exact name - return null to trigger error
                cubeField = null;
            }

            if (cubeField == null)
            {
                throw new InvalidOperationException($"Field '{fieldName}' not found in OLAP PivotTable. Use the exact CubeField name (e.g., '[Measures].[ACR]' or '[TableName].[ColumnName]').");
            }

            // CreatePivotFields() initializes PivotFields for fields not yet in the PivotTable.
            // It may throw if PivotFields already exist (field already in Values area).
            // Safe to ignore error - if PivotFields exist, we're good; if they don't and this fails,
            // subsequent operations will provide a more specific error.
            // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.cubefield.createpivotfields
            try
            {
                cubeField.CreatePivotFields();
            }
            catch
            {
                // PivotFields may already exist (field already added to PivotTable)
            }

            return cubeField; // Return CubeField, not PivotField
        }
        catch (Exception ex) when (cubeField == null)
        {
            throw new InvalidOperationException($"Field '{fieldName}' not found in OLAP PivotTable. Use the exact CubeField name (e.g., '[Measures].[ACR]' or '[TableName].[ColumnName]').", ex);
        }
        finally
        {
            ComUtilities.Release(ref cubeFields);
            // Note: Don't release cubeField - we're returning it
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldListResult ListFields(dynamic pivot, string workbookPath)
    {
        var fields = new List<PivotFieldInfo>();
        dynamic? cubeFields = null;

        try
        {
            cubeFields = pivot.CubeFields;
            int fieldCount = cubeFields.Count;

            for (int i = 1; i <= fieldCount; i++)
            {
                dynamic? cubeField = null;
                try
                {
                    cubeField = cubeFields.Item(i);
                    int orientation = Convert.ToInt32(cubeField.Orientation);

                    // Skip hidden fields
                    if (orientation == XlPivotFieldOrientation.xlHidden)
                        continue;

                    var fieldInfo = new PivotFieldInfo
                    {
                        Name = cubeField.Name?.ToString() ?? $"Field{i}",
                        CustomName = cubeField.Caption?.ToString() ?? "",
                        Area = (PivotFieldArea)orientation,
                        DataType = "Cube" // OLAP fields are always Cube type
                    };

                    // OLAP doesn't support AvailableValues like Regular PivotTables
                    // Values come from OLAP dimension hierarchies

                    fields.Add(fieldInfo);
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // Skip field if COM access fails - continue with other fields
                }
                finally
                {
                    ComUtilities.Release(ref cubeField);
                }
            }

            return new PivotFieldListResult
            {
                Success = true,
                Fields = fields,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeFields);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult AddRowField(dynamic pivot, string fieldName, int? position, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // Check if field is already placed
            int currentOrientation = Convert.ToInt32(cubeField.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first.");
            }

            // CRITICAL: Set Orientation on CubeField, NOT on PivotField
            cubeField.Orientation = XlPivotFieldOrientation.xlRowField;
            if (position.HasValue)
            {
                cubeField.Position = (double)position.Value;
            }

            // Refresh and validate
            pivot.RefreshTable();

            if (cubeField.Orientation != XlPivotFieldOrientation.xlRowField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Row area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Row,
                Position = Convert.ToInt32(cubeField.Position),
                DataType = "Cube",
                AvailableValues = new List<string>(), // OLAP doesn't provide unique values like Regular
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to add OLAP row field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult AddColumnField(dynamic pivot, string fieldName, int? position, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = Convert.ToInt32(cubeField.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first.");
            }

            cubeField.Orientation = XlPivotFieldOrientation.xlColumnField;
            if (position.HasValue)
            {
                cubeField.Position = (double)position.Value;
            }

            pivot.RefreshTable();

            if (cubeField.Orientation != XlPivotFieldOrientation.xlColumnField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Column area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Column,
                Position = Convert.ToInt32(cubeField.Position),
                DataType = "Cube",
                AvailableValues = new List<string>(),
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to add OLAP column field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult AddValueField(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string? customName, string workbookPath)
    {
        dynamic? cubeField = null;
        dynamic? workbook = null;
        dynamic? model = null;
        dynamic? modelTables = null;
        dynamic? table = null;
        dynamic? measures = null;
        dynamic? newMeasure = null;
        dynamic? formatObject;

        try
        {
            // TWO MODES:
            // MODE 1: Add pre-existing measure (fieldName starts with [Measures]. or already exists in Data Model)
            // MODE 2: Auto-create DAX measure from column (legacy behavior)

            // Get workbook and model
            workbook = pivot.Parent.Parent; // PivotTable -> Worksheet -> Workbook
            model = workbook.Model;

            if (model == null)
            {
                throw new InvalidOperationException(
                    $"Cannot add value field '{fieldName}' to OLAP PivotTable - workbook has no Data Model");
            }

            // MODE 1: Check if this is a pre-existing measure
            if (IsExistingMeasure(model, fieldName, out string? existingMeasureName))
            {
                // Find the measure's CubeField and add it to values area
                // IMPORTANT: Use exact match to avoid disambiguation bugs (e.g., "ACR" matching "ACRTypeKey")
                dynamic? cubeFields = null;
                try
                {
                    cubeFields = pivot.CubeFields;
                    for (int i = 1; i <= cubeFields.Count; i++)
                    {
                        dynamic? cf = null;
                        try
                        {
                            cf = cubeFields.Item(i);
                            string cfName = cf.Name?.ToString() ?? "";
                            int cubeFieldType = Convert.ToInt32(cf.CubeFieldType);

                            // Only match measures (CubeFieldType=2), not hierarchies (CubeFieldType=1)
                            // This prevents "ACR" from matching "[DisambiguationTable].[ACRTypeKey]"
                            if (cubeFieldType != XlCubeFieldType.xlMeasure)
                                continue;

                            // Check for exact match: [Measures].[MeasureName]
                            string expectedCubeFieldName = $"[Measures].[{existingMeasureName}]";
                            if (cfName.Equals(expectedCubeFieldName, StringComparison.OrdinalIgnoreCase) ||
                                cfName.Equals(existingMeasureName, StringComparison.OrdinalIgnoreCase) ||
                                cfName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                            {
                                cubeField = cf;
                                cf = null; // Transfer ownership
                                break;
                            }
                        }
                        finally
                        {
                            if (cf != null)
                                ComUtilities.Release(ref cf);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref cubeFields);
                }

                if (cubeField == null)
                {
                    throw new InvalidOperationException(
                        $"Measure '{existingMeasureName}' exists in Data Model but not found in PivotTable CubeFields. Try refreshing the PivotTable.");
                }

                // Check if measure is already in values area
                int currentOrientation = Convert.ToInt32(cubeField.Orientation);
                if (currentOrientation == XlPivotFieldOrientation.xlDataField)
                {
                    return new PivotFieldResult
                    {
                        Success = true,
                        FieldName = existingMeasureName!,
                        CustomName = cubeField.Caption?.ToString() ?? existingMeasureName!,
                        Area = PivotFieldArea.Value,
                        DataType = "Cube",
                        FilePath = workbookPath
                    };
                }

                // Add to values area using AddDataField (more reliable than setting Orientation directly)
                // Setting cubeField.Orientation = xlDataField can fail with E_INVALIDARG (0x80070057)
                // for CubeFields in certain states, while AddDataField works consistently
                pivot.AddDataField(cubeField);

                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = existingMeasureName!,
                    CustomName = customName ?? cubeField.Caption?.ToString() ?? existingMeasureName!,
                    Area = PivotFieldArea.Value,
                    Function = aggregationFunction,
                    DataType = "Cube",
                    FilePath = workbookPath
                };
            }

            // MODE 2: Create new measure from column (legacy auto-create behavior)
            // Find the source table and column for this field
            var tableAndColumn = FindTableAndColumn(pivot, fieldName);
            string tableName = tableAndColumn.Item1;
            string columnName = tableAndColumn.Item2;

            if (string.IsNullOrEmpty(tableName) || string.IsNullOrEmpty(columnName))
            {
                throw new InvalidOperationException(
                    $"Cannot determine table and column for field '{fieldName}'. " +
                    "Field must reference a Data Model table column (e.g., 'Sales' from 'SalesTable[Sales]') " +
                    "OR an existing measure (e.g., '[Measures].[Total Sales]')");
            }

            // Generate DAX formula and measure name
            string daxFormula = GenerateDaxFormula(tableName, columnName, aggregationFunction);
            string measureName = customName ?? $"{columnName} {GetFunctionName(aggregationFunction)}";

            // Find the table in the Data Model
            modelTables = model.ModelTables;
            for (int i = 1; i <= modelTables.Count; i++)
            {
                dynamic? t = null;
                try
                {
                    t = modelTables.Item(i);
                    string tName = t.Name?.ToString() ?? "";
                    if (tName.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        table = t;
                        t = null; // Transfer ownership
                        break;
                    }
                }
                finally
                {
                    if (t != null)
                        ComUtilities.Release(ref t);
                }
            }

            if (table == null)
            {
                throw new InvalidOperationException($"Table '{tableName}' not found in Data Model");
            }

            // Get ModelMeasures collection and create the measure
            measures = model.ModelMeasures;
            formatObject = GetDefaultFormatObject(model);

            newMeasure = measures.Add(
                measureName,
                table,
                daxFormula,
                formatObject,
                Type.Missing  // description
            );

            // Refresh the PivotTable connection to make the measure available in CubeFields
            pivot.RefreshTable();

            // Find the measure in CubeFields - measures appear with [Measures]. prefix
            // Use CubeFieldType to ensure we only match measures, not hierarchies
            dynamic? cubeFieldsForNewMeasure = null;
            try
            {
                cubeFieldsForNewMeasure = pivot.CubeFields;
                for (int i = 1; i <= cubeFieldsForNewMeasure.Count; i++)
                {
                    dynamic? cf = null;
                    try
                    {
                        cf = cubeFieldsForNewMeasure.Item(i);
                        string cfName = cf.Name?.ToString() ?? "";
                        int cubeFieldType = Convert.ToInt32(cf.CubeFieldType);

                        // Only match measures (CubeFieldType=2)
                        if (cubeFieldType != XlCubeFieldType.xlMeasure)
                            continue;

                        // Check for exact match: [Measures].[MeasureName]
                        string expectedCubeFieldName = $"[Measures].[{measureName}]";
                        if (cfName.Equals(expectedCubeFieldName, StringComparison.OrdinalIgnoreCase) ||
                            cfName.Equals(measureName, StringComparison.OrdinalIgnoreCase))
                        {
                            cubeField = cf;
                            cf = null; // Transfer ownership
                            break;
                        }
                    }
                    finally
                    {
                        if (cf != null)
                            ComUtilities.Release(ref cf);
                    }
                }
            }
            finally
            {
                ComUtilities.Release(ref cubeFieldsForNewMeasure);
            }

            if (cubeField == null)
            {
                throw new InvalidOperationException($"Measure '{measureName}' created but not found in PivotTable CubeFields after refresh");
            }

            // Add to values area using AddDataField (more reliable than setting Orientation directly)
            pivot.AddDataField(cubeField);

            return new PivotFieldResult
            {
                Success = true,
                FieldName = measureName,
                CustomName = customName ?? measureName,
                Area = PivotFieldArea.Value,
                Function = aggregationFunction,
                DataType = "Cube",
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = workbookPath
            };
        }
        finally
        {
            // Don't release formatObject - it's owned by the model
            ComUtilities.Release(ref newMeasure);
            ComUtilities.Release(ref measures);
            ComUtilities.Release(ref table);
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref workbook);
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult AddFilterField(dynamic pivot, string fieldName, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = Convert.ToInt32(cubeField.Orientation);
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {GetAreaName(currentOrientation)} area. Remove it first.");
            }

            cubeField.Orientation = XlPivotFieldOrientation.xlPageField;
            pivot.RefreshTable();

            if (cubeField.Orientation != XlPivotFieldOrientation.xlPageField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Filter area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Filter,
                Position = Convert.ToInt32(cubeField.Position),
                DataType = "Cube",
                AvailableValues = new List<string>(),
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to add OLAP filter field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult RemoveField(dynamic pivot, string fieldName, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = Convert.ToInt32(cubeField.Orientation);
            if (currentOrientation == XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is not currently placed in any area");
            }

            cubeField.Orientation = XlPivotFieldOrientation.xlHidden;
            pivot.RefreshTable();

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                Area = PivotFieldArea.Hidden,
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to remove OLAP field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult SetFieldName(dynamic pivot, string fieldName, string customName, string workbookPath)
    {
        dynamic? cubeField = null;
        try
        {
            // OLAP limitation: Cannot set Caption on CubeFields via COM
            throw new InvalidOperationException(
                $"Cannot rename OLAP field '{fieldName}' to '{customName}'. " +
                "Field names in OLAP PivotTables are derived from the Data Model definition. " +
                "To change field names: (1) Open Data Model in Excel, (2) Rename the dimension/hierarchy, (3) Refresh the PivotTable. " +
                "Reference: https://learn.microsoft.com/en-us/excel/vba/api/excel.cubefield.caption");
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult SetFieldFunction(dynamic pivot, string fieldName, AggregationFunction aggregationFunction, string workbookPath)
    {
        dynamic? workbook = null;
        dynamic? model = null;
        dynamic? measures = null;
        dynamic? measure = null;
        try
        {
            // For OLAP PivotTables, we need to update the DAX measure in the Data Model
            // Get workbook and model
            workbook = pivot.Parent.Parent;
            model = workbook.Model;

            if (model == null)
            {
                throw new InvalidOperationException(
                    $"Cannot update measure '{fieldName}' - workbook has no Data Model");
            }

            // Normalize field name - extract measure name from [Measures].[Name] format if present
            string targetMeasureName = NormalizeMeasureName(fieldName);

            // Find the measure by name
            measures = model.ModelMeasures;
            for (int i = 1; i <= measures.Count; i++)
            {
                dynamic? m = null;
                try
                {
                    m = measures.Item(i);
                    string mName = m.Name?.ToString() ?? "";
                    if (mName.Equals(targetMeasureName, StringComparison.OrdinalIgnoreCase))
                    {
                        measure = m;
                        m = null; // Transfer ownership
                        break;
                    }
                }
                finally
                {
                    if (m != null)
                        ComUtilities.Release(ref m);
                }
            }

            if (measure == null)
            {
                throw new InvalidOperationException($"Measure '{fieldName}' not found in Data Model");
            }

            // Parse the current formula to extract table and column
            string currentFormula = measure.Formula?.ToString() ?? "";
            var parsedFormula = ParseDaxFormula(currentFormula);

            if (string.IsNullOrEmpty(parsedFormula.tableName) || string.IsNullOrEmpty(parsedFormula.columnName))
            {
                throw new InvalidOperationException(
                    $"Cannot update measure '{fieldName}' - unable to parse current formula: {currentFormula}");
            }

            // Generate new DAX formula with the new aggregation function
            string newFormula = GenerateDaxFormula(parsedFormula.tableName, parsedFormula.columnName, aggregationFunction);

            // Update the measure's formula
            measure.Formula = newFormula;

            // Refresh the PivotTable to reflect the change
            pivot.RefreshTable();

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                Function = aggregationFunction,
                DataType = "Cube",
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref measure);
            ComUtilities.Release(ref measures);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref workbook);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult SetFieldFormat(dynamic pivot, string fieldName, string numberFormat, string workbookPath)
    {
        dynamic? cubeField = null;
        dynamic? pivotFields = null;
        dynamic? pivotField = null;
        try
        {
            // For OLAP PivotTables, find the CubeField and set NumberFormat on its PivotField
            // This works for measures in the Values area (including [Measures].[Name] format)
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // Verify the field is in the Values area (only data fields can have number formats)
            int orientation = Convert.ToInt32(cubeField.Orientation);
            if (orientation != XlPivotFieldOrientation.xlDataField)
            {
                throw new InvalidOperationException(
                    $"Field '{fieldName}' is not in the Values area (Orientation={orientation}). " +
                    "Only value fields can have number formats.");
            }

            // Access the PivotFields collection to set the NumberFormat
            // OLAP CubeFields expose their formatting through PivotFields
            pivotFields = cubeField.PivotFields;
            if (pivotFields == null || pivotFields.Count == 0)
            {
                throw new InvalidOperationException(
                    $"Cannot format OLAP field '{fieldName}' - PivotFields not available. " +
                    "Ensure the field has been added to the Values area.");
            }

            // Get the first (and typically only) PivotField and set its NumberFormat
            pivotField = pivotFields.Item(1);
            pivotField.NumberFormat = numberFormat;

            // NOTE: No RefreshTable() needed - NumberFormat is a visual-only property
            // RefreshTable() would re-query the Data Model which is very slow for OLAP PivotTables

            // Read back the format to verify it was set
            string? appliedFormat = null;
            try
            {
                appliedFormat = pivotField.NumberFormat?.ToString();
            }
            catch
            {
                // If we can't read it back, use what we set
                appliedFormat = numberFormat;
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Value,
                NumberFormat = appliedFormat,
                DataType = "Cube",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref pivotField);
            ComUtilities.Release(ref pivotFields);
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldFilterResult SetFieldFilter(dynamic pivot, string fieldName, List<string> filterValues, string workbookPath)
    {
        dynamic? cubeField = null;
        dynamic? pivotFields = null;
        dynamic? pivotField = null;
        dynamic? pivotItems = null;
        try
        {
            // OLAP limitation: Cannot set Visible property on OLAP PivotItems
            throw new InvalidOperationException(
                $"Cannot filter OLAP field '{fieldName}' via PivotItem.Visible property. " +
                "OLAP PivotItems do not support the Visible property. " +
                "To filter OLAP data: (1) Use PivotTable's built-in filter buttons in Excel, (2) Use OLAP Slicers for interactive filtering, or (3) Modify the source Data Model. " +
                "Reference: https://learn.microsoft.com/en-us/excel/vba/api/excel.pivotitem.visible");
        }
        catch (Exception ex)
        {
            return new PivotFieldFilterResult
            {
                Success = false,
                ErrorMessage = ex.Message,
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref pivotItems);
            ComUtilities.Release(ref pivotField);
            ComUtilities.Release(ref pivotFields);
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult SortField(dynamic pivot, string fieldName, SortDirection direction, string workbookPath)
    {
        dynamic? cubeField = null;
        dynamic? pivotFields = null;
        dynamic? pivotField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // OLAP sorting works through PivotField, not CubeField
            pivotFields = cubeField.PivotFields;
            if (pivotFields == null || pivotFields.Count == 0)
            {
                throw new InvalidOperationException($"Cannot sort OLAP field '{fieldName}' - PivotFields not available");
            }

            pivotField = pivotFields.Item(1);

            int sortOrder = direction == SortDirection.Ascending
                ? XlSortOrder.xlAscending
                : XlSortOrder.xlDescending;

            pivotField.AutoSort(sortOrder, fieldName);

            // NOTE: No RefreshTable() needed - Sorting is a visual-only operation

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = (PivotFieldArea)cubeField.Orientation,
                FilePath = workbookPath
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to sort OLAP field: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref pivotField);
            ComUtilities.Release(ref pivotFields);
            ComUtilities.Release(ref cubeField);
        }
    }

    /// <summary>
    /// Group a date/time field by the specified interval (Month, Quarter, Year).
    /// OLAP CubeFields automatically create date hierarchies from Data Model columns.
    /// Manual grouping via Group() is NOT supported for OLAP PivotTables.
    /// </summary>
    public PivotFieldResult GroupByDate(dynamic pivot, string fieldName, DateGroupingInterval interval, string workbookPath, ILogger? logger = null)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // OLAP PivotTables do not support manual date grouping via LabelRange.Group()
            // Date hierarchies are defined in the Data Model and automatically available
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Manual date grouping is not supported for OLAP PivotTables. " +
                              $"Date hierarchies must be defined in the Data Model. " +
                              $"Use Power Pivot to create date hierarchies (Year > Quarter > Month > Day) on the '{fieldName}' column.",
                FieldName = fieldName,
                FilePath = workbookPath,
                WorkflowHint = "For OLAP PivotTables: 1) Open Power Pivot, 2) Create date hierarchy on date column, " +
                               "3) Use RemoveField/AddField to place hierarchy levels in PivotTable areas."
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to access OLAP field '{fieldName}': {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult GroupByNumeric(dynamic pivot, string fieldName, double? start, double? endValue, double intervalSize, string workbookPath, ILogger? logger = null)
    {
        dynamic? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // OLAP PivotTables do not support manual numeric grouping via LabelRange.Group()
            // Numeric grouping must be done in the source data or Data Model
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Manual numeric grouping is not supported for OLAP PivotTables. " +
                              $"Numeric grouping must be defined in the Data Model. " +
                              $"Use Power Pivot to create calculated columns with range logic on the '{fieldName}' column.",
                FieldName = fieldName,
                FilePath = workbookPath,
                WorkflowHint = "For OLAP PivotTables: 1) Open Power Pivot, 2) Create calculated column with range logic " +
                               "(e.g., IF([Sales]<100, \"0-100\", IF([Sales]<200, \"100-200\", ...))), 3) Use that calculated column in PivotTable."
            };
        }
        catch (Exception ex)
        {
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Failed to access OLAP field '{fieldName}': {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref cubeField);
        }
    }

    /// <inheritdoc/>
    public PivotFieldResult CreateCalculatedField(dynamic pivot, string fieldName, string formula, string workbookPath, ILogger? logger = null)
    {
        // CRITICAL: OLAP PivotTables do NOT support CalculatedFields collection
        // The CalculatedFields collection returns Nothing for OLAP PivotTables
        // OLAP uses CalculatedMembers with MDX/DAX formulas instead
        //
        // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.calculatedfields
        // "For OLAP data sources, you cannot set this collection, and it always returns Nothing"
        return new PivotFieldResult
        {
            Success = false,
            FieldName = fieldName,
            Formula = formula,
            ErrorMessage = "Calculated fields are not supported for OLAP PivotTables. " +
                          "OLAP PivotTables use CalculatedMembers with MDX/DAX formulas instead. " +
                          "For Data Model PivotTables, use DAX measures via excel_datamodel tool.",
            FilePath = workbookPath,
            WorkflowHint = "For OLAP/Data Model PivotTables: " +
                          "1) Use excel_datamodel tool to create DAX measures with formulas, " +
                          "2) Refresh PivotTable to see new measures in field list, " +
                          "3) Add measure to Values area with AddValueField. " +
                          "Example DAX: Profit = SUM('Sales'[Revenue]) - SUM('Sales'[Cost])"
        };
    }

#pragma warning disable CA1848 // Keep logging for diagnostics
    /// <inheritdoc/>
    public OperationResult SetLayout(dynamic pivot, int layoutType, string workbookPath, ILogger? logger = null)
    {
        // OLAP PivotTables support all three layout forms
        // xlCompactRow=0, xlTabularRow=1, xlOutlineRow=2
        pivot.RowAxisLayout(layoutType);

        // NOTE: No RefreshTable() needed - Layout is a visual-only property

        if (logger?.IsEnabled(LogLevel.Information) == true)
        {
            logger.LogInformation("Set OLAP PivotTable layout to {LayoutType}", layoutType);
        }

        return new OperationResult
        {
            Success = true,
            FilePath = workbookPath
        };
    }
#pragma warning restore CA1848

#pragma warning disable CA1848 // Keep logging for diagnostics
    /// <inheritdoc/>
    public PivotFieldResult SetSubtotals(
        dynamic pivot,
        string fieldName,
        bool showSubtotals,
        string workbookPath,
        ILogger? logger = null)
    {
        dynamic? field = null;
        try
        {
            // Get the field - for OLAP, use PivotFields (not CubeFields)
            dynamic pivotFields = pivot.PivotFields;
            field = pivotFields.Item(fieldName);

            // OLAP PivotTables only support Automatic subtotals (index 1)
            // Other subtotal types not available for OLAP data sources
            field.Subtotals[1] = showSubtotals;

            // NOTE: No RefreshTable() needed - Subtotals is a visual-only property

            if (logger?.IsEnabled(LogLevel.Information) == true)
            {
                logger.LogInformation("Set OLAP subtotals for field {FieldName} to {ShowSubtotals}", fieldName, showSubtotals);
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                FilePath = workbookPath,
                WorkflowHint = showSubtotals
                    ? "Subtotals enabled for OLAP field. Only Automatic function supported (OLAP limitation)."
                    : "Subtotals disabled for OLAP field."
            };
        }
        catch (Exception ex)
        {
            if (logger?.IsEnabled(LogLevel.Error) == true)
            {
                logger.LogError(ex, "SetSubtotals failed for OLAP field {FieldName}", fieldName);
            }
            return new PivotFieldResult
            {
                Success = false,
                FieldName = fieldName,
                ErrorMessage = $"Failed to set OLAP subtotals: {ex.Message}",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref field);
        }
    }
#pragma warning restore CA1848

    /// <inheritdoc/>
#pragma warning disable CA1848
    public OperationResult SetGrandTotals(dynamic pivot, bool showRowGrandTotals, bool showColumnGrandTotals, string workbookPath, ILogger? logger = null)
    {
        pivot.RowGrand = showRowGrandTotals;
        pivot.ColumnGrand = showColumnGrandTotals;

        // NOTE: No RefreshTable() needed - GrandTotals are visual-only properties

        if (logger is not null && logger.IsEnabled(LogLevel.Information))
        {
            logger.LogInformation("Set OLAP grand totals: Row={RowGrand}, Column={ColumnGrand}", showRowGrandTotals, showColumnGrandTotals);
        }

        return new OperationResult
        {
            Success = true,
            FilePath = workbookPath
        };
    }
#pragma warning restore CA1848

    #region Helper Methods

    /// <summary>
    /// Find the source table and column for a CubeField in the Data Model.
    /// OLAP CubeFields reference Data Model columns in format: [TableName].[ColumnName]
    /// NOTE: This only searches hierarchy fields (CubeFieldType=1), not measures.
    /// </summary>
    private static (string tableName, string columnName) FindTableAndColumn(dynamic pivot, string fieldName)
    {
        dynamic? cubeFields = null;
        try
        {
            cubeFields = pivot.CubeFields;

            // Try to find the CubeField matching fieldName
            // Only look at hierarchies (table columns), not measures
            for (int i = 1; i <= cubeFields.Count; i++)
            {
                dynamic? cf = null;
                try
                {
                    cf = cubeFields.Item(i);
                    string cfName = cf.Name?.ToString() ?? "";
                    int cubeFieldType = Convert.ToInt32(cf.CubeFieldType);

                    // Skip measures - we're looking for table columns only
                    if (cubeFieldType == XlCubeFieldType.xlMeasure)
                        continue;

                    // EXACT MATCH ONLY - no partial matching
                    if (cfName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Parse hierarchical name format: [TableName].[ColumnName]
                        // Example: "[RegionalSalesTable].[Sales]" -> table="RegionalSalesTable", column="Sales"
                        if (cfName.Contains('[') && cfName.Contains(']'))
                        {
                            var parts = cfName.Split(FieldNameSeparators, StringSplitOptions.RemoveEmptyEntries);
                            if (parts.Length >= 2)
                            {
                                return (parts[0], parts[1]);
                            }
                        }

                        // Fallback: If no hierarchical format, assume fieldName is the column
                        // and try to infer table from the CubeField's SourceName property
                        try
                        {
                            string sourceName = cf.SourceName?.ToString() ?? "";
                            if (!string.IsNullOrEmpty(sourceName) && sourceName.Contains('['))
                            {
                                var sourceParts = sourceName.Split(FieldNameSeparators, StringSplitOptions.RemoveEmptyEntries);
                                if (sourceParts.Length >= 2)
                                {
                                    return (sourceParts[0], sourceParts[1]);
                                }
                            }
                        }
                        catch
                        {
                            // SourceName might not be available, continue with fallback
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref cf);
                }
            }

            return (string.Empty, string.Empty);
        }
        finally
        {
            ComUtilities.Release(ref cubeFields);
        }
    }

    /// <summary>
    /// Generate DAX formula for a measure based on aggregation function.
    /// Examples:
    /// - SUM: SUM('TableName'[ColumnName])
    /// - COUNT: COUNT('TableName'[ColumnName])
    /// - AVERAGE: AVERAGE('TableName'[ColumnName])
    /// </summary>
    private static string GenerateDaxFormula(string tableName, string columnName, AggregationFunction function)
    {
        string daxFunction = function switch
        {
            AggregationFunction.Sum => "SUM",
            AggregationFunction.Count => "COUNT",
            AggregationFunction.Average => "AVERAGE",
            AggregationFunction.Max => "MAX",
            AggregationFunction.Min => "MIN",
            AggregationFunction.CountNumbers => "COUNT",
            AggregationFunction.StdDev => "STDEV.S",
            AggregationFunction.StdDevP => "STDEV.P",
            AggregationFunction.Var => "VAR.S",
            AggregationFunction.VarP => "VAR.P",
            _ => throw new InvalidOperationException($"Unsupported aggregation function for DAX: {function}")
        };

        // DAX syntax: FUNCTION('TableName'[ColumnName])
        return $"{daxFunction}('{tableName}'[{columnName}])";
    }

    /// <summary>
    /// Get friendly function name for measure naming.
    /// </summary>
    private static string GetFunctionName(AggregationFunction function)
    {
        return function switch
        {
            AggregationFunction.Sum => "Sum",
            AggregationFunction.Count => "Count",
            AggregationFunction.Average => "Average",
            AggregationFunction.Max => "Max",
            AggregationFunction.Min => "Min",
            AggregationFunction.CountNumbers => "Count",
            AggregationFunction.StdDev => "StdDev",
            AggregationFunction.StdDevP => "StdDevP",
            AggregationFunction.Var => "Var",
            AggregationFunction.VarP => "VarP",
            _ => function.ToString()
        };
    }

    /// <summary>
    /// Get default format object from Data Model.
    /// Returns ModelFormatGeneral for standard numeric display.
    /// </summary>
    private static dynamic GetDefaultFormatObject(dynamic model)
    {
        // Get default format - ModelFormatGeneral is always available
        dynamic formats = model.ModelFormatGeneral;
        return formats;
    }

    /// <summary>
    /// Parse DAX formula to extract table and column names.
    /// Handles formats like: SUM('TableName'[ColumnName]), COUNT('Table'[Column]), etc.
    /// </summary>
    private static (string tableName, string columnName) ParseDaxFormula(string daxFormula)
    {
        // Expected format: FUNCTION('TableName'[ColumnName])
        // Extract table name from single quotes
        int tableStart = daxFormula.IndexOf('\'');
        int tableEnd = daxFormula.IndexOf('\'', tableStart + 1);

        if (tableStart == -1 || tableEnd == -1)
        {
            return (string.Empty, string.Empty);
        }

        string tableName = daxFormula.Substring(tableStart + 1, tableEnd - tableStart - 1);

        // Extract column name from square brackets
        int columnStart = daxFormula.IndexOf('[', tableEnd);
        int columnEnd = daxFormula.IndexOf(']', columnStart + 1);

        if (columnStart == -1 || columnEnd == -1)
        {
            return (string.Empty, string.Empty);
        }

        string columnName = daxFormula.Substring(columnStart + 1, columnEnd - columnStart - 1);

        return (tableName, columnName);
    }

    /// <summary>
    /// Parse number format string and create appropriate ModelFormat object.
    /// Supports: currency, percentage, decimal, whole number, general.
    /// </summary>
    private static dynamic? GetModelFormatObject(dynamic model, string numberFormat)
    {
        // Currency formats: $#,##0.00, $#,##0, etc.
        if (numberFormat.Contains('$'))
        {
            dynamic? currencyFormat = null;
            try
            {
                currencyFormat = model.ModelFormatCurrency;

                // Parse decimal places from format string
                int decimalIndex = numberFormat.IndexOf('.');
                if (decimalIndex >= 0)
                {
                    // Count zeros after decimal point
                    int decimalPlaces = 0;
                    for (int i = decimalIndex + 1; i < numberFormat.Length && numberFormat[i] == '0'; i++)
                    {
                        decimalPlaces++;
                    }
                    currencyFormat.DecimalPlaces = decimalPlaces;
                }
                else
                {
                    currencyFormat.DecimalPlaces = 0;
                }

                currencyFormat.Symbol = "$";
                return currencyFormat;
            }
            catch
            {
                if (currencyFormat != null)
                    ComUtilities.Release(ref currencyFormat);
                throw;
            }
        }

        // Percentage formats: 0.00%, 0%, etc.
        if (numberFormat.Contains('%'))
        {
            dynamic? percentFormat = null;
            try
            {
                percentFormat = model.ModelFormatPercentageNumber;

                // Parse decimal places
                int decimalIndex = numberFormat.IndexOf('.');
                if (decimalIndex >= 0)
                {
                    int decimalPlaces = 0;
                    for (int i = decimalIndex + 1; i < numberFormat.Length && numberFormat[i] == '0'; i++)
                    {
                        decimalPlaces++;
                    }
                    percentFormat.DecimalPlaces = decimalPlaces;
                }
                else
                {
                    percentFormat.DecimalPlaces = 0;
                }

                return percentFormat;
            }
            catch
            {
                if (percentFormat != null)
                    ComUtilities.Release(ref percentFormat);
                throw;
            }
        }

        // Decimal number formats: 0.00, #,##0.00, etc.
        if (numberFormat.Contains('.'))
        {
            dynamic? decimalFormat = null;
            try
            {
                decimalFormat = model.ModelFormatDecimalNumber;

                // Parse decimal places
                int decimalIndex = numberFormat.IndexOf('.');
                int decimalPlaces = 0;
                for (int i = decimalIndex + 1; i < numberFormat.Length && (numberFormat[i] == '0' || numberFormat[i] == '#'); i++)
                {
                    decimalPlaces++;
                }
                decimalFormat.DecimalPlaces = decimalPlaces;

                // Check for thousand separator
                if (numberFormat.Contains(','))
                {
                    decimalFormat.UseThousandSeparator = true;
                }

                return decimalFormat;
            }
            catch
            {
                if (decimalFormat != null)
                    ComUtilities.Release(ref decimalFormat);
                throw;
            }
        }

        // Whole number formats: 0, #,##0, etc.
        if (numberFormat.Contains('0') || numberFormat.Contains('#'))
        {
            dynamic? wholeFormat = null;
            try
            {
                wholeFormat = model.ModelFormatWholeNumber;

                // Check for thousand separator
                if (numberFormat.Contains(','))
                {
                    wholeFormat.UseThousandSeparator = true;
                }

                return wholeFormat;
            }
            catch
            {
                if (wholeFormat != null)
                    ComUtilities.Release(ref wholeFormat);
                throw;
            }
        }

        // Default: General format
        return model.ModelFormatGeneral;
    }

    /// <summary>
    /// Modify an existing format object's properties based on the format string.
    /// The format object is already attached to a measure and we modify it in place.
    /// </summary>
    private static void ModifyFormatObject(dynamic formatObject, string numberFormat)
    {
        // Try to determine the format type and modify accordingly
        // Currency format
        if (numberFormat.Contains('$'))
        {
            try
            {
                // Parse decimal places
                int decimalIndex = numberFormat.IndexOf('.');
                if (decimalIndex >= 0)
                {
                    int decimalPlaces = 0;
                    for (int i = decimalIndex + 1; i < numberFormat.Length && numberFormat[i] == '0'; i++)
                    {
                        decimalPlaces++;
                    }
                    formatObject.DecimalPlaces = decimalPlaces;
                }
                else
                {
                    formatObject.DecimalPlaces = 0;
                }

                formatObject.Symbol = "$";
                return;
            }
            catch
            {
                // If format object doesn't support these properties, it's not a currency format
                // Fall through to try other format types
            }
        }

        // Percentage format
        if (numberFormat.Contains('%'))
        {
            try
            {
                int decimalIndex = numberFormat.IndexOf('.');
                if (decimalIndex >= 0)
                {
                    int decimalPlaces = 0;
                    for (int i = decimalIndex + 1; i < numberFormat.Length && numberFormat[i] == '0'; i++)
                    {
                        decimalPlaces++;
                    }
                    formatObject.DecimalPlaces = decimalPlaces;
                }
                else
                {
                    formatObject.DecimalPlaces = 0;
                }
                return;
            }
            catch
            {
                // Not a percentage format
            }
        }

        // Decimal number format
        if (numberFormat.Contains('.'))
        {
            try
            {
                int decimalIndex = numberFormat.IndexOf('.');
                int decimalPlaces = 0;
                for (int i = decimalIndex + 1; i < numberFormat.Length && (numberFormat[i] == '0' || numberFormat[i] == '#'); i++)
                {
                    decimalPlaces++;
                }
                formatObject.DecimalPlaces = decimalPlaces;

                if (numberFormat.Contains(','))
                {
                    formatObject.UseThousandSeparator = true;
                }
                return;
            }
            catch
            {
                // Not a decimal format
            }
        }

        // Whole number format
        if (numberFormat.Contains('0') || numberFormat.Contains('#'))
        {
            try
            {
                if (numberFormat.Contains(','))
                {
                    formatObject.UseThousandSeparator = true;
                }
                return;
            }
            catch
            {
                // Not a whole number format
            }
        }

        // If we get here, it's probably ModelFormatGeneral which has no configurable properties
    }

    private static int GetComAggregationFunction(AggregationFunction function)
    {
        return function switch
        {
            AggregationFunction.Sum => XlConsolidationFunction.xlSum,
            AggregationFunction.Count => XlConsolidationFunction.xlCount,
            AggregationFunction.Average => XlConsolidationFunction.xlAverage,
            AggregationFunction.Max => XlConsolidationFunction.xlMax,
            AggregationFunction.Min => XlConsolidationFunction.xlMin,
            AggregationFunction.Product => XlConsolidationFunction.xlProduct,
            AggregationFunction.CountNumbers => XlConsolidationFunction.xlCountNums,
            AggregationFunction.StdDev => XlConsolidationFunction.xlStdDev,
            AggregationFunction.StdDevP => XlConsolidationFunction.xlStdDevP,
            AggregationFunction.Var => XlConsolidationFunction.xlVar,
            AggregationFunction.VarP => XlConsolidationFunction.xlVarP,
            _ => throw new InvalidOperationException($"Unsupported aggregation function: {function}")
        };
    }

    private static string GetAreaName(dynamic orientation)
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
    /// Normalize measure name by extracting it from [Measures].[Name] format if present.
    /// Returns the bare measure name (e.g., "Total Sales" from "[Measures].[Total Sales]").
    /// </summary>
    private static string NormalizeMeasureName(string fieldName)
    {
        if (fieldName.StartsWith("[Measures].", StringComparison.OrdinalIgnoreCase))
        {
            // Remove [Measures]. prefix and extract name from brackets
            var parts = fieldName.Split(FieldNameSeparators, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
            {
                return parts[1]; // "Total Sales" from "[Measures].[Total Sales]"
            }
        }
        return fieldName;
    }

    /// <summary>
    /// Check if fieldName refers to an existing measure in the Data Model.
    /// Returns true if the measure exists, and outputs the measure name.
    /// Handles formats: "[Measures].[Name]" or "Name" (exact match only).
    /// </summary>
    private static bool IsExistingMeasure(dynamic model, string fieldName, out string? measureName)
    {
        measureName = null;
        dynamic? measures = null;
        try
        {
            measures = model.ModelMeasures;
            if (measures == null || measures.Count == 0)
            {
                return false;
            }

            // Extract measure name from [Measures].[Name] format if present
            string searchName = NormalizeMeasureName(fieldName);

            // Search for measure by name (EXACT match only - no partial matching)
            // Partial matching causes disambiguation bugs where "ACR" could match "ACRTypeKey"
            for (int i = 1; i <= measures.Count; i++)
            {
                dynamic? measure = null;
                try
                {
                    measure = measures.Item(i);
                    string mName = measure.Name?.ToString() ?? "";

                    // Exact match only - no Contains() to avoid false positives
                    if (mName.Equals(searchName, StringComparison.OrdinalIgnoreCase) ||
                        mName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        measureName = mName;
                        return true;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref measure);
                }
            }

            return false;
        }
        finally
        {
            ComUtilities.Release(ref measures);
        }
    }

    #endregion
}

