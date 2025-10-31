using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Write operations - Delete, Create, Update
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public async Task<OperationResult> DeleteMeasureAsync(IExcelBatch batch, string measureName)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "model-delete-measure"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!ComInterop.ComUtilities.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModel();
                    return result;
                }

                model = ctx.Book.Model;

                // Find the measure
                measure = ComUtilities.FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    var measureNames = GetModelMeasureNames(model);
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.MeasureNotFound(measureName);

                    // Suggest similar measure names
                    var suggestions = new List<string>();
                    foreach (var m in measureNames)
                    {
                        if (m.Contains(measureName, StringComparison.OrdinalIgnoreCase))
                        {
                            suggestions.Add($"Try measure: {m}");
                            if (suggestions.Count >= 3) break;
                        }
                    }

                    if (suggestions.Any())
                    {
                        result.SuggestedNextActions = suggestions;
                    }

                    return result;
                }

                // Delete the measure
                measure.Delete();

                result.Success = true;
                result.SuggestedNextActions =
                [
                    $"Measure '{measureName}' deleted successfully",
                    "Use 'model-list-measures' to verify deletion",
                    "Changes saved to workbook"
                ];
                result.WorkflowHint = "Measure deleted. Next, verify remaining measures or create new ones.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("deleting measure", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref measure);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> DeleteRelationshipAsync(IExcelBatch batch, string fromTable, string fromColumn, string toTable, string toColumn)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "model-delete-relationship"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? modelRelationships = null;
            dynamic? relationship = null;
            try
            {
                // Check if workbook has Data Model
                if (!ComInterop.ComUtilities.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModel();
                    return result;
                }

                model = ctx.Book.Model;
                modelRelationships = model.ModelRelationships;

                // Find the relationship
                bool found = false;
                for (int i = 1; i <= modelRelationships.Count; i++)
                {
                    try
                    {
                        relationship = modelRelationships.Item(i);

                        dynamic? fkColumn = relationship.ForeignKeyColumn;
                        dynamic? pkColumn = relationship.PrimaryKeyColumn;

                        try
                        {
                            dynamic? fkTable = fkColumn.Parent;
                            dynamic? pkTable = pkColumn.Parent;

                            string currentFromTable = ComInterop.ComUtilities.SafeGetString(fkTable, "Name");
                            string currentFromColumn = ComInterop.ComUtilities.SafeGetString(fkColumn, "Name");
                            string currentToTable = ComInterop.ComUtilities.SafeGetString(pkTable, "Name");
                            string currentToColumn = ComInterop.ComUtilities.SafeGetString(pkColumn, "Name");

                            ComUtilities.Release(ref fkTable);
                            ComUtilities.Release(ref pkTable);

                            if (currentFromTable.Equals(fromTable, StringComparison.OrdinalIgnoreCase) &&
                                currentFromColumn.Equals(fromColumn, StringComparison.OrdinalIgnoreCase) &&
                                currentToTable.Equals(toTable, StringComparison.OrdinalIgnoreCase) &&
                                currentToColumn.Equals(toColumn, StringComparison.OrdinalIgnoreCase))
                            {
                                // Delete the relationship
                                relationship.Delete();
                                found = true;
                                break;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref fkColumn);
                            ComUtilities.Release(ref pkColumn);
                        }
                    }
                    finally
                    {
                        if (!found || i < modelRelationships.Count)
                        {
                            ComUtilities.Release(ref relationship);
                        }
                    }
                }

                if (!found)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.RelationshipNotFound(fromTable, fromColumn, toTable, toColumn);
                    result.SuggestedNextActions =
                    [
                        "Use 'model-list-relationships' to see available relationships",
                        "Check table and column names for typos",
                        "Verify the relationship exists in the Data Model"
                    ];
                    return result;
                }

                result.Success = true;
                result.SuggestedNextActions =
                [
                    $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} deleted successfully",
                    "Use 'model-list-relationships' to verify deletion",
                    "Changes saved to workbook"
                ];
                result.WorkflowHint = "Relationship deleted. Next, verify remaining relationships or create new ones.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("deleting relationship", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref relationship);
                ComUtilities.Release(ref modelRelationships);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateMeasureAsync(IExcelBatch batch, string tableName, string measureName,
                                                          string daxFormula, string? formatType = null,
                                                          string? description = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "model-create-measure"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? table = null;
            dynamic? measures = null;
            dynamic? newMeasure = null;
            dynamic? formatObject = null;
            try
            {
                // Check if workbook has Data Model
                if (!ComInterop.ComUtilities.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModel();
                    return result;
                }

                model = ctx.Book.Model;

                // Find the table
                table = ComUtilities.FindModelTable(model, tableName);
                if (table == null)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.TableNotFound(tableName);
                    return result;
                }

                // Check if measure already exists
                dynamic? existingMeasure = ComUtilities.FindModelMeasure(model, measureName);
                if (existingMeasure != null)
                {
                    ComUtilities.Release(ref existingMeasure);
                    result.Success = false;
                    result.ErrorMessage = $"Measure '{measureName}' already exists in the Data Model";
                    result.SuggestedNextActions =
                    [
                        "Use 'model-update-measure' to modify existing measure",
                        "Choose a different measure name",
                        "Delete the existing measure first"
                    ];
                    return result;
                }

                // Get ModelMeasures collection from MODEL (not from table!)
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.model.modelmeasures
                try
                {
                    measures = model.ModelMeasures;
                }
                catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                {
                    result.Success = false;
                    result.ErrorMessage = "DAX measures are not supported in this version of Excel. " +
                                        "The ModelMeasures API requires Microsoft Office 2016 or later. " +
                                        "Please upgrade Excel to use measure operations.";
                    return result;
                }

                // Get format object if specified
                if (!string.IsNullOrEmpty(formatType))
                {
                    formatObject = GetFormatObject(model, formatType);
                }

                // Create the measure using Excel COM API (Office 2016+)
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelmeasures.add
                newMeasure = measures.Add(
                    MeasureName: measureName,
                    AssociatedTable: table,
                    Formula: daxFormula,
                    FormatInformation: formatObject,
                    Description: description ?? ""
                );

                result.Success = true;
                result.SuggestedNextActions =
                [
                    $"Measure '{measureName}' created successfully in table '{tableName}'",
                    "Use 'model-view-measure' to verify the measure",
                    "Use 'model-list-measures' to see all measures",
                    "Changes saved to workbook"
                ];
                result.WorkflowHint = "Measure created. Next, test the measure in a PivotTable or verify its formula.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed($"creating measure '{measureName}'", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref formatObject);
                ComUtilities.Release(ref newMeasure);
                ComUtilities.Release(ref measures);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdateMeasureAsync(IExcelBatch batch, string measureName,
                                                          string? daxFormula = null, string? formatType = null,
                                                          string? description = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "model-update-measure"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            dynamic? formatObject = null;
            try
            {
                // Check if workbook has Data Model
                if (!ComInterop.ComUtilities.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModel();
                    return result;
                }

                model = ctx.Book.Model;

                // Find the measure
                measure = ComUtilities.FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.MeasureNotFound(measureName);
                    return result;
                }

                var updates = new List<string>();

                // Update formula if provided
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelmeasure (Formula property is Read/Write)
                if (!string.IsNullOrEmpty(daxFormula))
                {
                    measure.Formula = daxFormula;
                    updates.Add("Formula updated");
                }

                // Update format if provided
                if (!string.IsNullOrEmpty(formatType))
                {
                    formatObject = GetFormatObject(model, formatType);
                    if (formatObject != null)
                    {
                        measure.FormatInformation = formatObject;
                        updates.Add($"Format changed to {formatType}");
                    }
                }

                // Update description if provided
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelmeasure (Description property is Read/Write)
                if (description != null)
                {
                    measure.Description = description;
                    updates.Add("Description updated");
                }

                if (!updates.Any())
                {
                    result.Success = false;
                    result.ErrorMessage = "No updates provided. Specify at least one of: daxFormula, formatType, or description";
                    return result;
                }

                result.Success = true;
                result.SuggestedNextActions =
                [
                    $"Measure '{measureName}' updated: {string.Join(", ", updates)}",
                    "Use 'model-view-measure' to verify changes",
                    "Changes saved to workbook"
                ];
                result.WorkflowHint = "Measure updated. Next, test the changes in a PivotTable or verify the formula.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed($"updating measure '{measureName}'", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref formatObject);
                ComUtilities.Release(ref measure);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> CreateRelationshipAsync(IExcelBatch batch, string fromTable,
                                                                string fromColumn, string toTable,
                                                                string toColumn, bool active = true)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "model-create-relationship"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? relationships = null;
            dynamic? fromTableObj = null;
            dynamic? toTableObj = null;
            dynamic? fromColumnObj = null;
            dynamic? toColumnObj = null;
            dynamic? newRelationship = null;
            try
            {
                // Check if workbook has Data Model
                if (!ComInterop.ComUtilities.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModel();
                    return result;
                }

                model = ctx.Book.Model;

                // Find source table and column
                fromTableObj = ComUtilities.FindModelTable(model, fromTable);
                if (fromTableObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.TableNotFound(fromTable);
                    return result;
                }

                fromColumnObj = FindModelTableColumn(fromTableObj, fromColumn);
                if (fromColumnObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{fromColumn}' not found in table '{fromTable}'";
                    return result;
                }

                // Find target table and column
                toTableObj = ComUtilities.FindModelTable(model, toTable);
                if (toTableObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.TableNotFound(toTable);
                    return result;
                }

                toColumnObj = FindModelTableColumn(toTableObj, toColumn);
                if (toColumnObj == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Column '{toColumn}' not found in table '{toTable}'";
                    return result;
                }

                // Check if relationship already exists
                dynamic? existingRel = FindRelationship(model, fromTable, fromColumn, toTable, toColumn);
                if (existingRel != null)
                {
                    ComUtilities.Release(ref existingRel);
                    result.Success = false;
                    result.ErrorMessage = $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} already exists";
                    result.SuggestedNextActions =
                    [
                        "Use 'model-update-relationship' to modify relationship",
                        "Use 'model-list-relationships' to view all relationships"
                    ];
                    return result;
                }

                // Create the relationship using Excel COM API (Office 2016+)
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelrelationships.add
                relationships = model.ModelRelationships;
                newRelationship = relationships.Add(
                    ForeignKeyColumn: fromColumnObj,
                    PrimaryKeyColumn: toColumnObj
                );

                // Set active state
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelrelationship (Active property is Read/Write)
                newRelationship.Active = active;

                result.Success = true;
                result.SuggestedNextActions =
                [
                    $"Relationship created: {fromTable}.{fromColumn} → {toTable}.{toColumn} ({(active ? "Active" : "Inactive")})",
                    "Use 'model-list-relationships' to verify the relationship",
                    "Changes saved to workbook"
                ];
                result.WorkflowHint = "Relationship created. Next, test DAX calculations that use this relationship.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("creating relationship", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref newRelationship);
                ComUtilities.Release(ref toColumnObj);
                ComUtilities.Release(ref fromColumnObj);
                ComUtilities.Release(ref toTableObj);
                ComUtilities.Release(ref fromTableObj);
                ComUtilities.Release(ref relationships);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> UpdateRelationshipAsync(IExcelBatch batch, string fromTable,
                                                                string fromColumn, string toTable,
                                                                string toColumn, bool active)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "model-update-relationship"
        };

        return await batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? relationship = null;
            try
            {
                // Check if workbook has Data Model
                if (!ComInterop.ComUtilities.HasDataModel(ctx.Book))
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.NoDataModel();
                    return result;
                }

                model = ctx.Book.Model;

                // Find the relationship
                relationship = FindRelationship(model, fromTable, fromColumn, toTable, toColumn);
                if (relationship == null)
                {
                    result.Success = false;
                    result.ErrorMessage = DataModelErrorMessages.RelationshipNotFound(fromTable, fromColumn, toTable, toColumn);
                    result.SuggestedNextActions =
                    [
                        "Use 'model-list-relationships' to see available relationships",
                        "Check table and column names for typos"
                    ];
                    return result;
                }

                // Get current state
                bool wasActive = relationship.Active ?? false;

                // Update active state
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelrelationship (Active property is Read/Write)
                relationship.Active = active;

                string stateChange = wasActive == active
                    ? $"remains {(active ? "active" : "inactive")}"
                    : $"changed from {(wasActive ? "active" : "inactive")} to {(active ? "active" : "inactive")}";

                result.Success = true;
                result.SuggestedNextActions =
                [
                    $"Relationship {fromTable}.{fromColumn} → {toTable}.{toColumn} {stateChange}",
                    "Use 'model-list-relationships' to verify the change",
                    "Changes saved to workbook"
                ];
                result.WorkflowHint = "Relationship updated. Next, verify DAX calculations that use this relationship.";
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = DataModelErrorMessages.OperationFailed("updating relationship", ex.Message);
            }
            finally
            {
                ComUtilities.Release(ref relationship);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }
}
