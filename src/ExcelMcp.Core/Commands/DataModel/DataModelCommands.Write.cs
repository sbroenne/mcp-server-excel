using Polly;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Formatting;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.DataModel;
using Excel = Microsoft.Office.Interop.Excel;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Write operations - Delete, Create, Update
/// Includes resilient retry logic for intermittent 0x800AC472 errors.
/// </summary>
public partial class DataModelCommands
{
    // Resilience pipeline for Data Model operations - handles 0x800AC472 intermittent errors
    // See GitHub Issue #315: https://github.com/sbroenne/mcp-server-excel/issues/315
    private static readonly ResiliencePipeline _dataModelPipeline = ResiliencePipelines.CreateDataModelPipeline();

    /// <inheritdoc />
    public OperationResult DeleteMeasure(IExcelBatch batch, string measureName)
    {
        return ExecuteWithRetry(() =>
        {
            return batch.Execute((ctx, ct) =>
            {
                Excel.Model? model = null;
                Excel.ModelMeasure? measure = null;
                try
                {
                    // Check if workbook has Data Model
                    if (!HasDataModelTables(ctx.Book))
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                    }

                    model = ctx.Book.Model;

                    // Find the measure
                    measure = FindModelMeasure(model!, measureName);
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

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            });
        });
    }

    /// <inheritdoc />
    public OperationResult DeleteTable(IExcelBatch batch, string tableName)
    {
        return ExecuteWithRetry(() =>
        {
            return batch.Execute((ctx, ct) =>
            {
                Excel.Model? model = null;
                Excel.ModelTable? table = null;
                Excel.WorkbookConnection? sourceConnection = null;
                try
                {
                    // Check if workbook has Data Model
                    if (!HasDataModelTables(ctx.Book))
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                    }

                    model = ctx.Book.Model;

                    // Find the table
                    table = FindModelTable(model!, tableName);
                    if (table == null)
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.TableNotFound(tableName));
                    }

                    // IMPORTANT: ModelTable is read-only and cannot be deleted directly!
                    // The correct way to delete a Data Model table is to delete its
                    // SourceWorkbookConnection. When the connection is deleted,
                    // the associated ModelTable is automatically removed.
                    // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modeltable
                    // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.workbookconnection.delete
                    sourceConnection = table.SourceWorkbookConnection;
                    if (sourceConnection == null)
                    {
                        throw new InvalidOperationException(
                            $"Table '{tableName}' does not have an associated connection and cannot be deleted. " +
                            "This may indicate the table was created through an unsupported method.");
                    }

                    // Delete the connection, which removes the associated Data Model table
                    sourceConnection.Delete();
                }
                finally
                {
                    ComUtilities.Release(ref sourceConnection);
                    ComUtilities.Release(ref table);
                    ComUtilities.Release(ref model);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            });
        });
    }

    /// <inheritdoc />
    public OperationResult DeleteRelationship(IExcelBatch batch, string fromTable, string fromColumn, string toTable, string toColumn)
    {
        return ExecuteWithRetry(() =>
        {
            return batch.Execute((ctx, ct) =>
            {
                Excel.Model? model = null;
                Excel.ModelRelationships? modelRelationships = null;
                try
                {
                    // Check if workbook has Data Model
                    if (!HasDataModelTables(ctx.Book))
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                    }

                    model = ctx.Book.Model;
                    modelRelationships = model!.ModelRelationships;

                    // Find and delete the relationship
                    bool found = false;
                    int count = modelRelationships.Count;
                    for (int i = 1; i <= count; i++)
                    {
                        Excel.ModelRelationship? currentRelationship = null;
                        try
                        {
                            currentRelationship = modelRelationships.Item(i);

                            Excel.ModelTableColumn? fkColumn = currentRelationship.ForeignKeyColumn;
                            Excel.ModelTableColumn? pkColumn = currentRelationship.PrimaryKeyColumn;

                            try
                            {
                                Excel.ModelTable? fkTable = fkColumn?.Parent as Excel.ModelTable;
                                Excel.ModelTable? pkTable = pkColumn?.Parent as Excel.ModelTable;

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
                                    currentRelationship.Delete();
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
                            ComUtilities.Release(ref currentRelationship);
                        }
                    }

                    if (!found)
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.RelationshipNotFound(fromTable, fromColumn, toTable, toColumn));
                    }
                }
                finally
                {
                    ComUtilities.Release(ref modelRelationships);
                    ComUtilities.Release(ref model);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            });
        });
    }

    /// <inheritdoc />
    public OperationResult CreateMeasure(IExcelBatch batch, string tableName, string measureName,
                              string daxFormula, string? formatType = null,
                              string? description = null)
    {
        // Format DAX before saving (outside ExecuteWithRetry for async operation)
        // Formatting is done synchronously to maintain method signature compatibility
        // Falls back to original if formatting fails
        string formattedDax = DaxFormatter.FormatAsync(daxFormula).GetAwaiter().GetResult();

        return ExecuteWithRetry(() =>
        {
            return batch.Execute((ctx, ct) =>
            {
                Excel.Model? model = null;
                Excel.ModelTable? table = null;
                Excel.ModelMeasures? measures = null;
                Excel.ModelMeasure? newMeasure = null;
                object? formatObject = null;
                try
                {
                    // Check if workbook has Data Model
                    if (!HasDataModelTables(ctx.Book))
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                    }

                    model = ctx.Book.Model;

                    // Find the table
                    table = FindModelTable(model!, tableName);
                    if (table == null)
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.TableNotFound(tableName));
                    }

                    // Check if measure already exists
                    Excel.ModelMeasure? existingMeasure = FindModelMeasure(model!, measureName);
                    if (existingMeasure != null)
                    {
                        ComUtilities.Release(ref existingMeasure);
                        throw new InvalidOperationException($"Measure '{measureName}' already exists in the Data Model");
                    }

                    // Translate DAX formula separators from US format (comma) to locale-specific format
                    // This fixes issues on European locales where semicolon is the list separator
                    // Example: DATEADD(Date[Date], -1, MONTH) → DATEADD(Date[Date]; -1; MONTH) on German Excel
                    // NOTE: Translation is done on the FORMATTED DAX
                    var daxTranslator = new DaxFormulaTranslator(ctx.App);
                    string localizedFormula = daxTranslator.TranslateToLocale(formattedDax);

                    // Get ModelMeasures collection from MODEL (not from table!)
                    // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.model.modelmeasures
                    measures = model!.ModelMeasures;

                    // Get format object - ALWAYS returns a valid format object (never null)
                    // Fixed: Always provide format object to avoid failures on reopened Data Model files
                    formatObject = GetFormatObject(model!, formatType);

                    // Create the measure using Excel COM API (Office 2016+)
                    // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelmeasures.add
                    // FIXED: FormatInformation is REQUIRED (not optional as docs state)
                    // See: docs/KNOWN-ISSUES.md for details
                    newMeasure = measures.Add(
                        measureName,                                        // MeasureName (required)
                        table!,                                             // AssociatedTable (required)
                        localizedFormula,                                   // Formula (required) - must be valid DAX, formatted and translated for locale
                        formatObject!,                                      // FormatInformation (required) - NEVER null/Type.Missing
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

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            });
        });
    }

    /// <inheritdoc />
    public OperationResult UpdateMeasure(IExcelBatch batch, string measureName,
                              string? daxFormula = null, string? formatType = null,
                              string? description = null)
    {
        // Format DAX before saving (outside ExecuteWithRetry for async operation)
        // Only format if daxFormula is provided
        // Formatting is done synchronously to maintain method signature compatibility
        // Falls back to original if formatting fails
        string? formattedDax = null;
        if (!string.IsNullOrEmpty(daxFormula))
        {
            formattedDax = DaxFormatter.FormatAsync(daxFormula).GetAwaiter().GetResult();
        }

        return ExecuteWithRetry(() =>
        {
            return batch.Execute((ctx, ct) =>
            {
                Excel.Model? model = null;
                Excel.ModelMeasure? measure = null;
                object? formatObject = null;
                try
                {
                    // Check if workbook has Data Model
                    if (!HasDataModelTables(ctx.Book))
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                    }

                    model = ctx.Book.Model;

                    // Find the measure
                    measure = FindModelMeasure(model!, measureName);
                    if (measure == null)
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.MeasureNotFound(measureName));
                    }

                    var updates = new List<string>();

                    // Update formula if provided
                    // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelmeasure (Formula property is Read/Write)
                    if (!string.IsNullOrEmpty(formattedDax))
                    {
                        // Translate DAX formula separators from US format (comma) to locale-specific format
                        // This fixes issues on European locales where semicolon is the list separator
                        // NOTE: Translation is done on the FORMATTED DAX
                        var daxTranslator = new DaxFormulaTranslator(ctx.App);
                        string localizedFormula = daxTranslator.TranslateToLocale(formattedDax);
                        measure.Formula = localizedFormula;
                        updates.Add("Formula updated");
                    }

                    // Update format if provided
                    if (!string.IsNullOrEmpty(formatType))
                    {
                        formatObject = GetFormatObject(model!, formatType);
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

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            });
        });
    }

    /// <inheritdoc />
    public OperationResult CreateRelationship(IExcelBatch batch, string fromTable,
                                   string fromColumn, string toTable,
                                   string toColumn, bool active = true)
    {
        return ExecuteWithRetry(() =>
        {
            return batch.Execute((ctx, ct) =>
            {
                Excel.Model? model = null;
                Excel.ModelRelationships? relationships = null;
                Excel.ModelTable? fromTableObj = null;
                Excel.ModelTable? toTableObj = null;
                Excel.ModelTableColumn? fromColumnObj = null;
                Excel.ModelTableColumn? toColumnObj = null;
                Excel.ModelRelationship? newRelationship = null;
                try
                {
                    // Check if workbook has Data Model
                    if (!HasDataModelTables(ctx.Book))
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                    }

                    model = ctx.Book.Model;

                    // Find source table and column
                    fromTableObj = FindModelTable(model!, fromTable);
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
                    toTableObj = FindModelTable(model!, toTable);
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
                    Excel.ModelRelationship? existingRel = FindRelationship(model!, fromTable, fromColumn, toTable, toColumn);
                    if (existingRel != null)
                    {
                        ComUtilities.Release(ref existingRel);
                        throw new InvalidOperationException($"Relationship from {fromTable}.{fromColumn} to {toTable}.{toColumn} already exists");
                    }

                    // Create the relationship using Excel COM API (Office 2016+)
                    // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelrelationships.add
                    relationships = model!.ModelRelationships;
                    newRelationship = relationships.Add(
                        ForeignKeyColumn: fromColumnObj!,
                        PrimaryKeyColumn: toColumnObj!
                    );

                    // Set active state
                    // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelrelationship (Active property is Read/Write)
                    newRelationship!.Active = active;
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

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            });
        });
    }

    /// <inheritdoc />
    public OperationResult UpdateRelationship(IExcelBatch batch, string fromTable,
                                   string fromColumn, string toTable,
                                   string toColumn, bool active)
    {
        return ExecuteWithRetry(() =>
        {
            return batch.Execute((ctx, ct) =>
            {
                Excel.Model? model = null;
                Excel.ModelRelationship? relationship = null;
                try
                {
                    // Check if workbook has Data Model
                    if (!HasDataModelTables(ctx.Book))
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                    }

                    model = ctx.Book.Model;

                    // Find the relationship
                    relationship = FindRelationship(model!, fromTable, fromColumn, toTable, toColumn);
                    if (relationship == null)
                    {
                        throw new InvalidOperationException(DataModelErrorMessages.RelationshipNotFound(fromTable, fromColumn, toTable, toColumn));
                    }

                    // Update active state
                    // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.modelrelationship (Active property is Read/Write)
                    relationship.Active = active;
                }
                finally
                {
                    ComUtilities.Release(ref relationship);
                    ComUtilities.Release(ref model);
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            });
        });
    }

    /// <summary>
    /// Executes an action with resilient retry logic for intermittent Data Model errors.
    /// Handles 0x800AC472 and other transient COM errors with exponential backoff.
    /// </summary>
    /// <param name="action">The action to execute with retry</param>
    private static void ExecuteWithRetry(Action action)
    {
        _dataModelPipeline.Execute(action);
    }

    /// <summary>
    /// Executes a function with resilient retry logic for intermittent Data Model errors.
    /// Returns the result of the function.
    /// </summary>
    /// <typeparam name="T">The return type</typeparam>
    /// <param name="func">The function to execute with retry</param>
    private static T ExecuteWithRetry<T>(Func<T> func)
    {
        return _dataModelPipeline.Execute(func);
    }
}



