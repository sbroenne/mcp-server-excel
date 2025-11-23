using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Write operations - Delete, Create, Update
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public void DeleteMeasure(IExcelBatch batch, string measureName)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Find the measure
                measure = FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    throw new InvalidOperationException(DataModelErrorMessages.MeasureNotFound(measureName));
                }

                // Delete the measure
                measure.Delete();
            }
            finally
            {
                ComUtilities.Release(ref measure);
                ComUtilities.Release(ref model);
            }

            return 0;
        });
    }

    /// <inheritdoc />
    public void DeleteRelationship(IExcelBatch batch, string fromTable, string fromColumn, string toTable, string toColumn)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? modelRelationships = null;
            dynamic? relationship = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
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

                            string currentFromTable = ComUtilities.SafeGetString(fkTable, "Name");
                            string currentFromColumn = ComUtilities.SafeGetString(fkColumn, "Name");
                            string currentToTable = ComUtilities.SafeGetString(pkTable, "Name");
                            string currentToColumn = ComUtilities.SafeGetString(pkColumn, "Name");

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
                    throw new InvalidOperationException(DataModelErrorMessages.RelationshipNotFound(fromTable, fromColumn, toTable, toColumn));
                }
            }
            finally
            {
                ComUtilities.Release(ref relationship);
                ComUtilities.Release(ref modelRelationships);
                ComUtilities.Release(ref model);
            }

            return 0;
        });
    }

    /// <inheritdoc />
    public void CreateMeasure(IExcelBatch batch, string tableName, string measureName,
                              string daxFormula, string? formatType = null,
                              string? description = null)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? table = null;
            dynamic? measures = null;
            dynamic? newMeasure = null;
            dynamic? formatObject = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Find the table
                table = FindModelTable(model, tableName);
                if (table == null)
                {
                    throw new InvalidOperationException(DataModelErrorMessages.TableNotFound(tableName));
                }

                // Check if measure already exists
                dynamic? existingMeasure = FindModelMeasure(model, measureName);
                if (existingMeasure != null)
                {
                    ComUtilities.Release(ref existingMeasure);
                    throw new InvalidOperationException($"Measure '{measureName}' already exists in the Data Model");
                }

                // Get ModelMeasures collection from MODEL (not from table!)
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.model.modelmeasures
                measures = model.ModelMeasures;

                // Get format object - ALWAYS returns a valid format object (never null)
                // Fixed: Always provide format object to avoid failures on reopened Data Model files
                formatObject = GetFormatObject(model, formatType);

                // Create the measure using Excel COM API (Office 2016+)
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelmeasures.add
                // FIXED: FormatInformation is REQUIRED (not optional as docs state)
                // See: docs/KNOWN-ISSUES.md for details
                newMeasure = measures.Add(
                    measureName,                                        // MeasureName (required)
                    table,                                              // AssociatedTable (required)
                    daxFormula,                                         // Formula (required) - must be valid DAX
                    formatObject,                                       // FormatInformation (required) - NEVER null/Type.Missing
                    string.IsNullOrEmpty(description) ? Type.Missing : description  // Description (optional)
                );
            }
            finally
            {
                // Note: formatObject is a property reference from the model (not a new object)
                // Do NOT release formatObject - it's owned by the model
                ComUtilities.Release(ref newMeasure);
                ComUtilities.Release(ref measures);
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref model);
            }

            return 0;
        });
    }

    /// <inheritdoc />
    public void UpdateMeasure(IExcelBatch batch, string measureName,
                              string? daxFormula = null, string? formatType = null,
                              string? description = null)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? measure = null;
            dynamic? formatObject = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Find the measure
                measure = FindModelMeasure(model, measureName);
                if (measure == null)
                {
                    throw new InvalidOperationException(DataModelErrorMessages.MeasureNotFound(measureName));
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

                if (updates.Count == 0)
                {
                    throw new ArgumentException("No updates provided. Specify at least one of: daxFormula, formatType, or description");
                }
            }
            finally
            {
                // Note: formatObject is a property reference from the model (not a new object)
                // Do NOT release formatObject - it's owned by the model
                ComUtilities.Release(ref measure);
                ComUtilities.Release(ref model);
            }

            return 0;
        });
    }

    /// <inheritdoc />
    public void CreateRelationship(IExcelBatch batch, string fromTable,
                                   string fromColumn, string toTable,
                                   string toColumn, bool active = true)
    {
        batch.Execute((ctx, ct) =>
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
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Find source table and column
                fromTableObj = FindModelTable(model, fromTable);
                if (fromTableObj == null)
                {
                    throw new InvalidOperationException(DataModelErrorMessages.TableNotFound(fromTable));
                }

                fromColumnObj = FindModelTableColumn(fromTableObj, fromColumn);
                if (fromColumnObj == null)
                {
                    throw new InvalidOperationException($"Column '{fromColumn}' not found in table '{fromTable}'");
                }

                // Find target table and column
                toTableObj = FindModelTable(model, toTable);
                if (toTableObj == null)
                {
                    throw new InvalidOperationException(DataModelErrorMessages.TableNotFound(toTable));
                }

                toColumnObj = FindModelTableColumn(toTableObj, toColumn);
                if (toColumnObj == null)
                {
                    throw new InvalidOperationException($"Column '{toColumn}' not found in table '{toTable}'");
                }

                // Check if relationship already exists
                dynamic? existingRel = FindRelationship(model, fromTable, fromColumn, toTable, toColumn);
                if (existingRel != null)
                {
                    ComUtilities.Release(ref existingRel);
                    throw new InvalidOperationException($"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} already exists");
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

            return 0;
        });
    }

    /// <inheritdoc />
    public void UpdateRelationship(IExcelBatch batch, string fromTable,
                                   string fromColumn, string toTable,
                                   string toColumn, bool active)
    {
        batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? relationship = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Find the relationship
                relationship = FindRelationship(model, fromTable, fromColumn, toTable, toColumn);
                if (relationship == null)
                {
                    throw new InvalidOperationException(DataModelErrorMessages.RelationshipNotFound(fromTable, fromColumn, toTable, toColumn));
                }

                // Get current state
                bool wasActive = relationship.Active ?? false;

                // Update active state
                // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelrelationship (Active property is Read/Write)
                relationship.Active = active;

                string stateChange = wasActive == active
                    ? $"remains {(active ? "active" : "inactive")}"
                    : $"changed from {(wasActive ? "active" : "inactive")} to {(active ? "active" : "inactive")}";
            }
            finally
            {
                ComUtilities.Release(ref relationship);
                ComUtilities.Release(ref model);
            }

            return 0;
        });
    }
}

