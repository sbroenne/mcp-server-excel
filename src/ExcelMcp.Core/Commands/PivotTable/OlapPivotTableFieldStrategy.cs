using Microsoft.Extensions.Logging;
using System.Text.RegularExpressions;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands.PivotTable;

/// <summary>
/// Strategy for OLAP (Online Analytical Processing) PivotTable field operations.
/// Uses CubeFields API for Data Model-based PivotTables.
///
/// Unplaced OLAP hierarchies may need CreatePivotFields() before their fields
/// can be manipulated. Measures and materialized fields do not need this call.
/// Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.cubefield.createpivotfields
/// </summary>
public partial class OlapPivotTableFieldStrategy : IPivotTableFieldStrategy
{
    private static Excel.CubeField GetFieldForManipulation(Excel.PivotTable pivot, string fieldName)
    {
        Excel.CubeFields? cubeFields = null;
        Excel.CubeField? cubeField = null;
        Excel.PivotFields? existingFields = null;
        bool ownershipTransferred = false;
        try
        {
            cubeFields = pivot.CubeFields;

            // EXACT MATCH ONLY - no partial matching to avoid disambiguation bugs
            // The LLM knows exact field names, so partial matching only causes problems
            // (e.g., "ACR" incorrectly matching "[DisambiguationTable].[ACRTypeKey]")
            try
            {
                cubeField = cubeFields[fieldName];
            }
            catch (System.Runtime.InteropServices.COMException ex) when (ComErrors.IsMemberUnavailable(ex))
            {
                // Field not found by exact name - return null to trigger error
                cubeField = null;
            }

            if (cubeField == null)
            {
                throw new InvalidOperationException($"Field '{fieldName}' not found in OLAP PivotTable. Use the exact CubeField name (e.g., '[Measures].[ACR]' or '[TableName].[ColumnName]').");
            }

            // CreatePivotFields initializes a hierarchy before it is placed. Measures and
            // already materialized hierarchies do not need (and can reject) this call.
            if ((int)cubeField.CubeFieldType == (int)Excel.XlCubeFieldType.xlHierarchy &&
                cubeField.Orientation == Excel.XlPivotFieldOrientation.xlHidden)
            {
                existingFields = (Excel.PivotFields)cubeField.PivotFields;
                if (existingFields.Count == 0)
                    cubeField.CreatePivotFields();
            }

            ownershipTransferred = true;
            return cubeField; // Return CubeField, not PivotField
        }
        catch (Exception ex) when (cubeField is null && CreatedObjectFailure.CanReport(ex))
        {
            throw new InvalidOperationException($"Field '{fieldName}' not found in OLAP PivotTable. Use the exact CubeField name (e.g., '[Measures].[ACR]' or '[TableName].[ColumnName]').", ex);
        }
        finally
        {
            ComUtilities.Release(ref existingFields);
            if (!ownershipTransferred)
                ComUtilities.Release(ref cubeField);
            ComUtilities.Release(ref cubeFields);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldListResult ListFields(Excel.PivotTable pivot, string workbookPath)
    {
        var fields = new List<PivotFieldInfo>();
        Excel.CubeFields? cubeFields = null;

        try
        {
            cubeFields = pivot.CubeFields;
            int fieldCount = cubeFields.Count;

            for (int i = 1; i <= fieldCount; i++)
            {
                Excel.CubeField? cubeField = null;
                try
                {
                    cubeField = cubeFields[i];
                    int orientation = (int)cubeField.Orientation;

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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Field-operation failure transport returns Success=false with the original native message; cancellation and disconnected sessions propagate.")]
    public PivotFieldResult AddRowField(Excel.PivotTable pivot, string fieldName, int? position, string workbookPath)
    {
        Excel.CubeField? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // Check if field is already placed
            int currentOrientation = (int)cubeField.Orientation;
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            // CRITICAL: Set Orientation on CubeField, NOT on PivotField
            cubeField.Orientation = Excel.XlPivotFieldOrientation.xlRowField;
            if (position.HasValue)
            {
                cubeField.Position = position.Value;
            }

            // NOTE: No RefreshTable() needed - orientation change takes effect immediately
            // RefreshTable() causes RPC disconnection on rapid operations (issue #426)

            if (cubeField.Orientation != Excel.XlPivotFieldOrientation.xlRowField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Row area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Row,
                Position = Convert.ToInt32((object)cubeField.Position, System.Globalization.CultureInfo.InvariantCulture),
                DataType = "Cube",
                AvailableValues = [],
                FilePath = workbookPath
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Column-field failure transport returns Success=false with the original native message; cancellation and disconnected sessions propagate.")]
    public PivotFieldResult AddColumnField(Excel.PivotTable pivot, string fieldName, int? position, string workbookPath)
    {
        Excel.CubeField? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = (int)cubeField.Orientation;
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            cubeField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField;
            if (position.HasValue)
            {
                cubeField.Position = position.Value;
            }

            // NOTE: No RefreshTable() needed - orientation change takes effect immediately
            // RefreshTable() causes RPC disconnection on rapid operations (issue #426)

            if (cubeField.Orientation != Excel.XlPivotFieldOrientation.xlColumnField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Column area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Column,
                Position = Convert.ToInt32((object)cubeField.Position, System.Globalization.CultureInfo.InvariantCulture),
                DataType = "Cube",
                AvailableValues = [],
                FilePath = workbookPath
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Value-field failure transport reports Success=false and the original error; cancellation and disconnected sessions propagate without pretending the operation completed.")]
    public PivotFieldResult AddValueField(Excel.PivotTable pivot, string fieldName, AggregationFunction aggregationFunction, string? customName, string workbookPath)
    {
        Excel.CubeField? cubeField = null;
        Excel.Worksheet? worksheet = null;
        Excel.Workbook? workbook = null;
        Excel.Model? model = null;
        Excel.ModelTables? modelTables = null;
        Excel.ModelTable? table = null;
        Excel.ModelMeasures? measures = null;
        Excel.ModelMeasure? newMeasure = null;
        object? formatObject = null;
        Excel.PivotField? dataField = null;
        Excel.PivotFields? pivotFields = null;

        try
        {
            // TWO MODES:
            // MODE 1: Add a measure the PivotTable already exposes (external cube or Data Model)
            // MODE 2: Auto-create DAX measure from a Data Model column (legacy behavior)

            // External cube measures exist only on the server, never in the workbook Data Model.
            string? existingMeasureName = null;
            cubeField = FindMeasureCubeField(pivot, fieldName);
            if (cubeField == null)
            {
                if (!UsesWorkbookDataModel(pivot))
                {
                    throw new InvalidOperationException(
                        $"Field '{fieldName}' is not a measure of the connected external cube. " +
                        "Measures must be defined on the cube server; workbook DAX measures cannot be created for this PivotTable.");
                }
                worksheet = (Excel.Worksheet)pivot.Parent;
                workbook = (Excel.Workbook)worksheet.Parent;
                model = workbook.Model;
                if (model != null && IsExistingMeasure(model, fieldName, out existingMeasureName))
                {
                    cubeField = FindMeasureCubeField(pivot, existingMeasureName!);
                    if (cubeField == null)
                        throw new InvalidOperationException(
                            $"Measure '{existingMeasureName}' exists in Data Model but not found in PivotTable CubeFields. Try refreshing the PivotTable.");
                }
            }

            if (cubeField != null)
            {
                existingMeasureName = GetBareMeasureName(cubeField.Name?.ToString() ?? fieldName);
                // Check if measure is already in values area
                int currentOrientation = (int)cubeField.Orientation;
                if (currentOrientation == XlPivotFieldOrientation.xlDataField)
                {
                    pivotFields = cubeField.PivotFields;
                    dataField = pivotFields.Item(1);
                    if (customName != null)
                        dataField.Caption = customName;
                    return new PivotFieldResult
                    {
                        Success = true,
                        FieldName = existingMeasureName!,
                        CustomName = dataField.Caption,
                        Area = PivotFieldArea.Value,
                        DataType = "Cube",
                        FilePath = workbookPath
                    };
                }

                // Add to values area using AddDataField (more reliable than setting Orientation directly)
                // Setting cubeField.Orientation = xlDataField can fail with E_INVALIDARG (0x80070057)
                // for CubeFields in certain states, while AddDataField works consistently
                dataField = pivot.AddDataField(cubeField, customName ?? Type.Missing);

                return new PivotFieldResult
                {
                    Success = true,
                    FieldName = existingMeasureName!,
                    CustomName = dataField.Caption,
                    Area = PivotFieldArea.Value,
                    DataType = "Cube",
                    FilePath = workbookPath
                };
            }

            // MODE 2: Create new measure from column (legacy auto-create behavior)
            if (model == null)
            {
                throw new InvalidOperationException(
                    $"Cannot add value field '{fieldName}' to OLAP PivotTable - it is not a measure of the connected cube and the workbook has no Data Model");
            }

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
                Excel.ModelTable? t = null;
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
            cubeField = FindMeasureCubeField(pivot, measureName);

            if (cubeField == null)
            {
                throw new InvalidOperationException($"Measure '{measureName}' created but not found in PivotTable CubeFields after refresh");
            }

            // Add to values area using AddDataField (more reliable than setting Orientation directly)
            dataField = pivot.AddDataField(cubeField);

            return new PivotFieldResult
            {
                Success = true,
                FieldName = measureName,
                CustomName = dataField.Caption,
                Area = PivotFieldArea.Value,
                Function = aggregationFunction,
                DataType = "Cube",
                FilePath = workbookPath
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
            ComUtilities.Release(ref dataField);
            ComUtilities.Release(ref pivotFields);
            ComUtilities.Release(ref newMeasure);
            ComUtilities.Release(ref formatObject);
            ComUtilities.Release(ref measures);
            ComUtilities.Release(ref table);
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
            ComUtilities.Release(ref workbook);
            ComUtilities.Release(ref worksheet);
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Filter-field failure transport returns Success=false with the original native message; cancellation and disconnected sessions propagate.")]
    public PivotFieldResult AddFilterField(Excel.PivotTable pivot, string fieldName, string workbookPath)
    {
        Excel.CubeField? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = (int)cubeField.Orientation;
            if (currentOrientation != XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is already placed in {PivotTableHelpers.GetAreaName(currentOrientation)} area. Remove it first.");
            }

            cubeField.Orientation = Excel.XlPivotFieldOrientation.xlPageField;

            // NOTE: No RefreshTable() needed - orientation change takes effect immediately
            // RefreshTable() causes RPC disconnection on rapid operations (issue #426)

            if (cubeField.Orientation != Excel.XlPivotFieldOrientation.xlPageField)
            {
                throw new InvalidOperationException($"Field '{fieldName}' was not successfully added to Filter area.");
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption?.ToString() ?? fieldName,
                Area = PivotFieldArea.Filter,
                Position = Convert.ToInt32((object)cubeField.Position, System.Globalization.CultureInfo.InvariantCulture),
                DataType = "Cube",
                AvailableValues = [],
                FilePath = workbookPath
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Removal failure transport returns Success=false and the original error; it performs no rollback, and cancellation/disconnection propagate.")]
    public PivotFieldResult RemoveField(Excel.PivotTable pivot, string fieldName, string workbookPath)
    {
        Excel.CubeField? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            int currentOrientation = (int)cubeField.Orientation;
            if (currentOrientation == XlPivotFieldOrientation.xlHidden)
            {
                throw new InvalidOperationException($"Field '{fieldName}' is not currently placed in any area");
            }

            cubeField.Orientation = Excel.XlPivotFieldOrientation.xlHidden;

            // NOTE: No RefreshTable() needed - orientation change takes effect immediately
            // RefreshTable() causes RPC disconnection on rapid operations (issue #426)

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                Area = PivotFieldArea.Hidden,
                FilePath = workbookPath
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Reports this unsupported OLAP rename as Success=false with the original NotSupportedException message.")]
    public PivotFieldResult SetFieldName(Excel.PivotTable pivot, string fieldName, string customName, string workbookPath)
    {
        Excel.CubeField? cubeField = null;
        try
        {
            // OLAP limitation: Cannot set Caption on CubeFields via COM
            throw new InvalidOperationException(
                $"Cannot rename OLAP field '{fieldName}' to '{customName}'. " +
                "Field names in OLAP PivotTables come from the Data Model definition or, for an external cube server, from the cube itself. " +
                "To change field names: (1) Rename the dimension/hierarchy in the Data Model, or on the cube server, (2) Refresh the PivotTable. " +
                "Reference: https://learn.microsoft.com/en-us/excel/vba/api/excel.cubefield.caption");
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Aggregation failure transport reports Success=false and the actual error; cancellation and disconnected sessions are not converted to results.")]
    public PivotFieldResult SetFieldFunction(Excel.PivotTable pivot, string fieldName, AggregationFunction aggregationFunction, string workbookPath)
    {
        Excel.Worksheet? worksheet = null;
        Excel.Workbook? workbook = null;
        Excel.Model? model = null;
        Excel.ModelMeasures? measures = null;
        Excel.ModelMeasure? measure = null;
        try
        {
            if (!UsesWorkbookDataModel(pivot))
                throw new InvalidOperationException(
                    $"Cannot update measure '{fieldName}' - an external cube measure is defined on the cube server and must be changed there.");
            // For OLAP PivotTables, we need to update the DAX measure in the Data Model
            // Get workbook and model
            worksheet = (Excel.Worksheet)pivot.Parent;
            workbook = (Excel.Workbook)worksheet.Parent;
            model = workbook.Model;

            if (model == null)
            {
                throw new InvalidOperationException(
                    $"Cannot update measure '{fieldName}' - workbook has no Data Model. A measure from an external cube server is defined on the server and must be changed there.");
            }

            // Normalize field name - extract measure name from [Measures].[Name] format if present
            string targetMeasureName = GetBareMeasureName(fieldName);

            // Find the measure by name
            measures = model.ModelMeasures;
            for (int i = 1; i <= measures.Count; i++)
            {
                Excel.ModelMeasure? m = null;
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
                throw new InvalidOperationException(
                    $"Measure '{fieldName}' not found in the workbook Data Model. " +
                    "Only Data Model measures can be changed here; a measure from an external cube server is defined on the server and must be changed there.");
            }

            // Parse the current formula to extract table and column
            string currentFormula = measure.Formula?.ToString() ?? "";
            var parsedFormula = ParseDaxFormula(currentFormula);

            if (string.IsNullOrEmpty(parsedFormula.tableName) || string.IsNullOrEmpty(parsedFormula.columnName))
            {
                throw new InvalidOperationException(
                    $"Cannot update measure '{fieldName}' - only a single aggregation of one Data Model column can be changed here. " +
                    "Use datamodel update-measure to change a compound DAX formula.");
            }

            // Generate new DAX formula with the new aggregation function
            string newFormula = GenerateDaxFormula(parsedFormula.tableName, parsedFormula.columnName, aggregationFunction);

            // Update the measure's formula
            measure.Formula = newFormula;

            // NOTE: No RefreshTable() needed - formula change takes effect immediately
            // RefreshTable() causes RPC disconnection on rapid operations (issue #426)

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                Function = aggregationFunction,
                DataType = "Cube",
                FilePath = workbookPath
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
            ComUtilities.Release(ref worksheet);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    public PivotFieldResult SetFieldFormat(Excel.PivotTable pivot, string fieldName, string numberFormat, string workbookPath)
    {
        Excel.CubeField? cubeField = null;
        Excel.PivotFields? dataFields = null;
        Excel.PivotField? dataField = null;
        Excel.Range? dataRange = null;
        Excel.Range? cells = null;
        Excel.Range? firstDataCell = null;
        try
        {
            cubeField = FindMeasureCubeField(pivot, fieldName);
            if (cubeField is not null && cubeField.Orientation != Excel.XlPivotFieldOrientation.xlDataField)
            {
                throw new InvalidOperationException(
                    $"Field '{fieldName}' is not in the Values area (Orientation={cubeField.Orientation}). " +
                    "Only value fields can have number formats.");
            }

            // CubeField.PivotFields exposes cube metadata, not the displayed data field.
            dataFields = (Excel.PivotFields)((Excel.PivotTable)pivot).DataFields;
            string? cubeFieldName = cubeField?.Name;
            for (int index = 1; index <= dataFields.Count; index++)
            {
                Excel.PivotField? candidate = null;
                Excel.CubeField? candidateCubeField = null;
                try
                {
                    candidate = dataFields.Item(index);
                    candidateCubeField = candidate.CubeField;
                    bool matches = cubeFieldName is not null
                        ? string.Equals(candidateCubeField.Name, cubeFieldName, StringComparison.OrdinalIgnoreCase)
                        : string.Equals(candidate.Caption, fieldName, StringComparison.OrdinalIgnoreCase) ||
                          string.Equals(candidateCubeField.Caption, fieldName, StringComparison.OrdinalIgnoreCase);
                    if (matches)
                    {
                        if (dataField is not null)
                        {
                            throw new InvalidOperationException(
                                $"Measure caption '{fieldName}' is ambiguous in the OLAP PivotTable. Use its exact CubeField name.");
                        }

                        dataField = candidate;
                        candidate = null;
                    }
                }
                finally
                {
                    ComUtilities.Release(ref candidateCubeField);
                    ComUtilities.Release(ref candidate);
                }
            }

            if (dataField is null)
            {
                throw new InvalidOperationException(
                    $"Cannot format OLAP measure '{fieldName}' because its data field was not found. " +
                    "Use its exact CubeField name or displayed caption and ensure it is in the Values area.");
            }

            dataField.NumberFormat = numberFormat;
            string? appliedFormat = dataField.NumberFormat;
            string? renderedText = null;
            dataRange = dataField.DataRange;
            if (dataRange is not null)
            {
                cells = dataRange.Cells;
                firstDataCell = (Excel.Range)cells[1, 1];
                renderedText = ((object?)firstDataCell.Text)?.ToString();
            }

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = dataField.Caption,
                Area = PivotFieldArea.Value,
                NumberFormat = appliedFormat,
                RenderedText = renderedText,
                DataType = "Cube",
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref firstDataCell);
            ComUtilities.Release(ref cells);
            ComUtilities.Release(ref dataRange);
            ComUtilities.Release(ref dataField);
            ComUtilities.Release(ref dataFields);
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>

    /// <inheritdoc/>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Reports unsupported OLAP item filtering as Success=false with the actual NotSupportedException message.")]
    public PivotFieldFilterResult SetFieldFilter(Excel.PivotTable pivot, string fieldName, List<string> filterValues, string workbookPath)
    {
        Excel.CubeField? cubeField = null;
        Excel.PivotFields? pivotFields = null;
        Excel.PivotField? pivotField = null;
        Excel.PivotItems? pivotItems = null;
        try
        {
            // OLAP limitation: Cannot set Visible property on OLAP PivotItems
            throw new InvalidOperationException(
                $"Cannot filter OLAP field '{fieldName}' via PivotItem.Visible property. " +
                "OLAP PivotItems do not support the Visible property. " +
                "To filter OLAP data: (1) Use PivotTable's built-in filter buttons in Excel, (2) Use OLAP Slicers for interactive filtering, or (3) Modify the source Data Model. " +
                "Reference: https://learn.microsoft.com/en-us/excel/vba/api/excel.pivotitem.visible");
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Sort failure transport returns Success=false with the actual error; cancellation and disconnected sessions propagate.")]
    public PivotFieldResult SortField(Excel.PivotTable pivot, string fieldName, SortDirection direction, string workbookPath)
    {
        Excel.CubeField? cubeField = null;
        Excel.PivotFields? pivotFields = null;
        Excel.PivotField? pivotField = null;
        try
        {
            cubeField = (Excel.CubeField)GetFieldForManipulation(pivot, fieldName);

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

            // Excel needs the level's unique name, not the enclosing CubeField hierarchy.
            pivotField.AutoSort(sortOrder, pivotField.SourceName);

            // NOTE: No RefreshTable() needed - Sorting is a visual-only operation

            return new PivotFieldResult
            {
                Success = true,
                FieldName = fieldName,
                CustomName = cubeField.Caption ?? fieldName,
                Area = (PivotFieldArea)cubeField.Orientation,
                FilePath = workbookPath
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Grouping prerequisite failure transport returns Success=false and the real access error; cancellation and disconnection propagate.")]
    public PivotFieldResult GroupByDate(Excel.PivotTable pivot, string fieldName, DateGroupingInterval interval, string workbookPath, ILogger? logger = null)
    {
        Excel.CubeField? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // OLAP PivotTables do not support manual date grouping via LabelRange.Group()
            // Date hierarchies are defined in the Data Model and automatically available
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Manual date grouping is not supported for OLAP PivotTables. " +
                              $"Date hierarchies must be defined in the Data Model, or on the cube server for an external cube. " +
                              $"Use Power Pivot (or the cube's own design) to create date hierarchies (Year > Quarter > Month > Day) for '{fieldName}'.",
                FieldName = fieldName,
                FilePath = workbookPath,
                WorkflowHint = "For a workbook Data Model, create a date hierarchy in Power Pivot. " +
                               "For an external cube, use an existing date hierarchy or have it defined on the cube server. " +
                               "Then use RemoveField/AddField to place its levels in PivotTable areas."
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Numeric-grouping prerequisite failure transport returns Success=false and the real access error; cancellation and disconnection propagate.")]
    public PivotFieldResult GroupByNumeric(Excel.PivotTable pivot, string fieldName, double? start, double? endValue, double intervalSize, string workbookPath, ILogger? logger = null)
    {
        Excel.CubeField? cubeField = null;
        try
        {
            cubeField = GetFieldForManipulation(pivot, fieldName);

            // OLAP PivotTables do not support manual numeric grouping via LabelRange.Group()
            // Numeric grouping must be done in the source data or Data Model
            return new PivotFieldResult
            {
                Success = false,
                ErrorMessage = $"Manual numeric grouping is not supported for OLAP PivotTables. " +
                              $"Numeric grouping must be defined in the Data Model, or on the cube server for an external cube. " +
                              $"Use Power Pivot (or the cube's own design) to create range logic for '{fieldName}'.",
                FieldName = fieldName,
                FilePath = workbookPath,
                WorkflowHint = "For a workbook Data Model, create a calculated column with range logic in Power Pivot. " +
                               "For an external cube, use an existing grouping field or have it defined on the cube server. " +
                               "Then use that field in the PivotTable."
            };
        }
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
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
    public PivotFieldResult CreateCalculatedField(Excel.PivotTable pivot, string fieldName, string formula, string workbookPath, ILogger? logger = null)
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
                          "For Data Model PivotTables, use DAX measures via datamodel tool.",
            FilePath = workbookPath,
            WorkflowHint = "For OLAP/Data Model PivotTables: " +
                          "1) Use datamodel tool to create DAX measures with formulas, " +
                          "2) Refresh PivotTable to see new measures in field list, " +
                          "3) Add measure to Values area with AddValueField. " +
                          "Example DAX: Profit = SUM('Sales'[Revenue]) - SUM('Sales'[Cost])"
        };
    }
    /// <inheritdoc/>
    public OperationResult SetLayout(Excel.PivotTable pivot, int rowLayout, string workbookPath, ILogger? logger = null)
    {
        // OLAP PivotTables support all three layout forms
        // xlCompactRow=0, xlTabularRow=1, xlOutlineRow=2
        pivot.RowAxisLayout((Excel.XlLayoutRowType)rowLayout);

        // NOTE: No RefreshTable() needed - Layout is a visual-only property

        if (logger?.IsEnabled(LogLevel.Information) is true)
        {
            logger.LogInformation("Set OLAP PivotTable layout to {LayoutType}", rowLayout);
        }

        return new OperationResult
        {
            Success = true,
            FilePath = workbookPath
        };
    }
    /// <inheritdoc/>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002", Justification = "Partial subtotal failure transport returns Success=false and describes the real error, completed hierarchy levels and uncertain attempted level without rollback; cancellation and disconnection propagate.")]
    public PivotFieldResult SetSubtotals(
        Excel.PivotTable pivot,
        string fieldName,
        bool showSubtotals,
        string workbookPath,
        ILogger? logger = null)
    {
        Excel.CubeField? cubeField = null;
        Excel.PivotFields? pivotFields = null;
        var completedLevels = new List<string>();
        string? attemptedLevel = null;
        try
        {
            cubeField = (Excel.CubeField)GetFieldForManipulation(pivot, fieldName);
            if (cubeField.Orientation is not (Excel.XlPivotFieldOrientation.xlRowField or Excel.XlPivotFieldOrientation.xlColumnField))
                throw new InvalidOperationException("Subtotals require an OLAP hierarchy in the Row or Column area.");
            pivotFields = cubeField.PivotFields;
            for (int index = 1; index <= pivotFields.Count; index++)
            {
                Excel.PivotField? field = null;
                try
                {
                    attemptedLevel = $"level {index}";
                    field = pivotFields.Item(index);
                    attemptedLevel = field.Name;
                    field.Subtotals[1] = showSubtotals;
                    completedLevels.Add(attemptedLevel);
                    attemptedLevel = null;
                }
                finally { ComUtilities.Release(ref field); }
            }

            // NOTE: No RefreshTable() needed - Subtotals is a visual-only property

            if (logger?.IsEnabled(LogLevel.Information) is true)
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
        catch (Exception ex) when (CreatedObjectFailure.CanReport(ex))
        {
            if (logger?.IsEnabled(LogLevel.Error) is true)
            {
                logger.LogError(ex, "SetSubtotals failed for OLAP field {FieldName}", fieldName);
            }
            return new PivotFieldResult
            {
                Success = false,
                FieldName = fieldName,
                ErrorMessage = DescribeSubtotalFailure(ex, completedLevels, attemptedLevel),
                FilePath = workbookPath
            };
        }
        finally
        {
            ComUtilities.Release(ref pivotFields);
            ComUtilities.Release(ref cubeField);
        }
    }
    /// <inheritdoc/>
    public OperationResult SetGrandTotals(Excel.PivotTable pivot, bool showRowGrandTotals, bool showColumnGrandTotals, string workbookPath, ILogger? logger = null)
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

    #region Helper Methods

    internal static string DescribeSubtotalFailure(Exception error, IReadOnlyList<string> completedLevels, string? attemptedLevel)
    {
        return $"Failed to set OLAP subtotals: {error.Message}. " +
            $"Levels updated: {(completedLevels.Count == 0 ? "none" : string.Join(", ", completedLevels))}. " +
            $"Attempted level: {attemptedLevel ?? "none"}" +
            (attemptedLevel is null ? ". " : "; it may also have changed before the failure. ") +
            "No changes have been rolled back. Inspect the hierarchy before retrying.";
    }

    private static bool UsesWorkbookDataModel(Excel.PivotTable pivot)
    {
        Excel.PivotCache? cache = null;
        Excel.WorkbookConnection? connection = null;
        try
        {
            cache = ((Excel.PivotTable)pivot).PivotCache();
            connection = cache.WorkbookConnection;
            return connection?.Type == Excel.XlConnectionType.xlConnectionTypeMODEL;
        }
        finally
        {
            ComUtilities.Release(ref connection);
            ComUtilities.Release(ref cache);
        }
    }

    /// <summary>
    /// Find the source table and column for a CubeField in the Data Model.
    /// OLAP CubeFields reference Data Model columns in format: [TableName].[ColumnName]
    /// NOTE: This only searches hierarchy fields (CubeFieldType=1), not measures.
    /// </summary>
    private static (string tableName, string columnName) FindTableAndColumn(Excel.PivotTable pivot, string fieldName)
    {
        Excel.CubeFields? cubeFields = null;
        try
        {
            cubeFields = pivot.CubeFields;

            // Try to find the CubeField matching fieldName
            // Only look at hierarchies (table columns), not measures
            for (int i = 1; i <= cubeFields.Count; i++)
            {
                Excel.CubeField? cf = null;
                try
                {
                    cf = cubeFields[i];
                    string cfName = cf.Name?.ToString() ?? "";
                    int cubeFieldType = (int)cf.CubeFieldType;

                    // Skip measures - we're looking for table columns only
                    if (cubeFieldType == XlCubeFieldType.xlMeasure)
                        continue;

                    // EXACT MATCH ONLY - no partial matching
                    if (cfName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    {
                        var match = DataModelColumnName().Match(cfName);
                        if (match.Success)
                        {
                            return (match.Groups["table"].Value.Replace("]]", "]", StringComparison.Ordinal),
                                match.Groups["column"].Value.Replace("]]", "]", StringComparison.Ordinal));
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
        return $"{daxFunction}('{tableName.Replace("'", "''", StringComparison.Ordinal)}'[{columnName.Replace("]", "]]", StringComparison.Ordinal)}])";
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
    private static object GetDefaultFormatObject(Excel.Model model)
    {
        // Get default format - ModelFormatGeneral is always available
        return model.ModelFormatGeneral;
    }

    /// <summary>
    /// Parse DAX formula to extract table and column names.
    /// Handles formats like: SUM('TableName'[ColumnName]), COUNT('Table'[Column]), etc.
    /// </summary>
    internal static (string tableName, string columnName) ParseDaxFormula(string daxFormula)
    {
        var match = SingleAggregationFormula().Match(daxFormula);
        return match.Success
            ? (match.Groups["table"].Value.Replace("''", "'", StringComparison.Ordinal),
                match.Groups["column"].Value.Replace("]]", "]", StringComparison.Ordinal))
            : (string.Empty, string.Empty);
    }

    [GeneratedRegex(@"^\[(?<table>(?:[^\]]|\]\])+)\]\.\[(?<column>(?:[^\]]|\]\])+)\]$",
        RegexOptions.CultureInvariant, 1000)]
    private static partial Regex DataModelColumnName();

    [GeneratedRegex(@"^\s*=?\s*(?:SUM|COUNT|AVERAGE|MAX|MIN|STDEV\.S|STDEV\.P|VAR\.S|VAR\.P)\s*\(\s*'(?<table>(?:[^']|'')+)'\s*\[(?<column>(?:[^\]]|\]\])+)\]\s*\)\s*$",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant, 1000)]
    private static partial Regex SingleAggregationFormula();

    /// <summary>
    /// Finds a measure CubeField by exact name, accepting "[Measures].[Name]" or bare "Name".
    /// Only measures match, so "ACR" never matches a hierarchy such as "[Table].[ACRTypeKey]".
    /// Caller owns the returned COM object.
    /// </summary>
    private static Excel.CubeField? FindMeasureCubeField(Excel.PivotTable pivot, string name)
    {
        string qualified = GetQualifiedMeasureName(name);
        Excel.CubeFields? cubeFields = null;
        try
        {
            cubeFields = pivot.CubeFields;
            for (int i = 1; i <= cubeFields.Count; i++)
            {
                Excel.CubeField? cf = null;
                try
                {
                    cf = cubeFields[i];
                    if ((int)cf.CubeFieldType != XlCubeFieldType.xlMeasure)
                        continue;

                    string cfName = cf.Name?.ToString() ?? "";
                    if (cfName.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                        cfName.Equals(qualified, StringComparison.OrdinalIgnoreCase))
                    {
                        var found = cf;
                        cf = null; // Transfer ownership
                        return found;
                    }
                }
                finally
                {
                    if (cf != null)
                        ComUtilities.Release(ref cf);
                }
            }
            return null;
        }
        finally
        {
            ComUtilities.Release(ref cubeFields);
        }
    }

    // "[Measures].[Amount.SUM]" -> "Amount.SUM"; dots inside a measure name are preserved.
    internal static string GetBareMeasureName(string cubeFieldName)
    {
        const string prefix = "[Measures].[";
        return cubeFieldName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) && cubeFieldName.EndsWith(']')
            ? cubeFieldName[prefix.Length..^1].Replace("]]", "]", StringComparison.Ordinal)
            : cubeFieldName;
    }

    internal static string GetQualifiedMeasureName(string name) =>
        $"[Measures].[{GetBareMeasureName(name).Replace("]", "]]", StringComparison.Ordinal)}]";

    /// <summary>
    /// Check if fieldName refers to an existing measure in the Data Model.
    /// Returns true if the measure exists, and outputs the measure name.
    /// Handles formats: "[Measures].[Name]" or "Name" (exact match only).
    /// </summary>
    private static bool IsExistingMeasure(Excel.Model model, string fieldName, out string? measureName)
    {
        measureName = null;
        Excel.ModelMeasures? measures = null;
        try
        {
            measures = model.ModelMeasures;
            if (measures is null || measures.Count == 0)
            {
                return false;
            }

            // Extract measure name from [Measures].[Name] format if present
            string searchName = GetBareMeasureName(fieldName);

            // Search for measure by name (EXACT match only - no partial matching)
            // Partial matching causes disambiguation bugs where "ACR" could match "ACRTypeKey"
            for (int i = 1; i <= measures!.Count; i++)
            {
                Excel.ModelMeasure? measure = null;
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
