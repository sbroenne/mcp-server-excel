using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.ComInterop.Session;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

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

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
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
                    var measureNames = DataModelHelpers.GetModelMeasureNames(model);
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
                result.SuggestedNextActions = new List<string>
                {
                    $"Measure '{measureName}' deleted successfully",
                    "Use 'model-list-measures' to verify deletion",
                    "Changes saved to workbook"
                };
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

        return await batch.ExecuteAsync(async (ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? modelRelationships = null;
            dynamic? relationship = null;
            try
            {
                // Check if workbook has Data Model
                if (!DataModelHelpers.HasDataModel(ctx.Book))
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

                            string currentFromTable = DataModelHelpers.SafeGetString(fkTable, "Name");
                            string currentFromColumn = DataModelHelpers.SafeGetString(fkColumn, "Name");
                            string currentToTable = DataModelHelpers.SafeGetString(pkTable, "Name");
                            string currentToColumn = DataModelHelpers.SafeGetString(pkColumn, "Name");

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
                    result.SuggestedNextActions = new List<string>
                    {
                        "Use 'model-list-relationships' to see available relationships",
                        "Check table and column names for typos",
                        "Verify the relationship exists in the Data Model"
                    };
                    return result;
                }

                result.Success = true;
                result.SuggestedNextActions = new List<string>
                {
                    $"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} deleted successfully",
                    "Use 'model-list-relationships' to verify deletion",
                    "Changes saved to workbook"
                };
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

    // Phase 2: CREATE/UPDATE methods will be added here:
    // - CreateMeasureAsync
    // - UpdateMeasureAsync
    // - CreateRelationshipAsync
    // - UpdateRelationshipAsync
}
