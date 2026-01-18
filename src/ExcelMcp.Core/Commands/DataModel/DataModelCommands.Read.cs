using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model Read operations - List, View, Export
/// </summary>
public partial class DataModelCommands
{
    /// <inheritdoc />
    public DataModelTableListResult ListTables(IExcelBatch batch)
    {
        var result = new DataModelTableListResult { FilePath = batch.WorkbookPath };

        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(2));

        return batch.Execute((ctx, ct) =>
        {
            // Check if workbook has Data Model
            if (!HasDataModelTables(ctx.Book))
            {
                // Empty Data Model is valid - return empty list (LLM-friendly)
                result.Success = true;
                result.Tables = [];
                return result;
            }

            dynamic? model = null;
            try
            {
                model = ctx.Book.Model;

                ForEachTable(model, (Action<dynamic, int>)((table, index) =>
                {
                    var tableInfo = new DataModelTableInfo
                    {
                        Name = ComUtilities.SafeGetString(table, "Name"),
                        SourceName = ComUtilities.SafeGetString(table, "SourceName"),
                        RecordCount = ComUtilities.SafeGetInt(table, "RecordCount")
                    };

                    result.Tables.Add(tableInfo);
                }));

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        }, timeoutCts.Token);
    }

    /// <inheritdoc />
    public DataModelMeasureListResult ListMeasures(IExcelBatch batch, string? tableName = null)
    {
        var result = new DataModelMeasureListResult { FilePath = batch.WorkbookPath };

        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(2));

        result = batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Iterate through all measures (they're at model level)
                ForEachMeasure(model, (Action<dynamic, int>)((measure, index) =>
                {
                    // Get the table name for this measure
                    string measureTableName = string.Empty;
                    dynamic? associatedTable = null;
                    try
                    {
                        associatedTable = measure.AssociatedTable;
                        measureTableName = associatedTable?.Name?.ToString() ?? string.Empty;
                    }
                    finally
                    {
                        ComUtilities.Release(ref associatedTable);
                    }

                    // Skip if filtering by table and this measure isn't in that table
                    if (tableName != null && !measureTableName.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    string formula = ComUtilities.SafeGetString(measure, "Formula");

                    var measureInfo = new DataModelMeasureInfo
                    {
                        Name = ComUtilities.SafeGetString(measure, "Name"),
                        Table = measureTableName,
                        FormulaPreview = formula, // Will be formatted after retrieval
                        Description = ComUtilities.SafeGetString(measure, "Description")
                    };

                    result.Measures.Add(measureInfo);
                }));

                // Check if table filter was specified but not found
                if (tableName != null && result.Measures.Count == 0)
                {
                    throw new InvalidOperationException(DataModelErrorMessages.TableNotFound(tableName));
                }

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        }, timeoutCts.Token);

        return result;
    }

    /// <inheritdoc />
    public DataModelMeasureViewResult Read(IExcelBatch batch, string measureName)
    {
        var result = new DataModelMeasureViewResult
        {
            FilePath = batch.WorkbookPath,
            MeasureName = measureName
        };

        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(2));

        result = batch.Execute((ctx, ct) =>
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

                // Get measure details using safe helpers
                result.DaxFormula = ComUtilities.SafeGetString(measure, "Formula");
                result.Description = ComUtilities.SafeGetString(measure, "Description");
                result.CharacterCount = result.DaxFormula.Length;
                result.TableName = GetMeasureTableName(model, measureName) ?? "";

                // Try to get format information - FormatInformation returns ModelFormat* objects
                // (ModelFormatGeneral, ModelFormatCurrency, ModelFormatDecimalNumber, etc.)
                // These don't have a FormatString property - they have type-specific properties
                dynamic? formatInfo = null;
                try
                {
                    formatInfo = measure.FormatInformation;
                    if (formatInfo != null)
                    {
                        result.FormatInfo = GetFormatInfo(formatInfo);
                    }
                }
                finally
                {
                    ComUtilities.Release(ref formatInfo);
                }

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref measure);
                ComUtilities.Release(ref model);
            }

            return result;
        });

        return result;
    }

    /// <inheritdoc />
    public DataModelRelationshipListResult ListRelationships(IExcelBatch batch)
    {
        var result = new DataModelRelationshipListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                ForEachRelationship(model, (Action<dynamic, int>)((relationship, index) =>
                {
                    var relInfo = new DataModelRelationshipInfo
                    {
                        FromTable = ComUtilities.SafeGetString(relationship.ForeignKeyColumn?.Parent, "Name"),
                        FromColumn = ComUtilities.SafeGetString(relationship.ForeignKeyColumn, "Name"),
                        ToTable = ComUtilities.SafeGetString(relationship.PrimaryKeyColumn?.Parent, "Name"),
                        ToColumn = ComUtilities.SafeGetString(relationship.PrimaryKeyColumn, "Name"),
                        IsActive = relationship.Active ?? false
                    };

                    result.Relationships.Add(relInfo);
                }));

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelRelationshipViewResult ReadRelationship(IExcelBatch batch, string fromTable, string fromColumn, string toTable, string toColumn)
    {
        var result = new DataModelRelationshipViewResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Find the matching relationship
                DataModelRelationshipInfo? foundRelationship = null;
                ForEachRelationship(model, (Action<dynamic, int>)((relationship, index) =>
                {
                    var relFromTable = ComUtilities.SafeGetString(relationship.ForeignKeyColumn?.Parent, "Name");
                    var relFromColumn = ComUtilities.SafeGetString(relationship.ForeignKeyColumn, "Name");
                    var relToTable = ComUtilities.SafeGetString(relationship.PrimaryKeyColumn?.Parent, "Name");
                    var relToColumn = ComUtilities.SafeGetString(relationship.PrimaryKeyColumn, "Name");

                    if (string.Equals(relFromTable, fromTable, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(relFromColumn, fromColumn, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(relToTable, toTable, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(relToColumn, toColumn, StringComparison.OrdinalIgnoreCase))
                    {
                        foundRelationship = new DataModelRelationshipInfo
                        {
                            FromTable = relFromTable,
                            FromColumn = relFromColumn,
                            ToTable = relToTable,
                            ToColumn = relToColumn,
                            IsActive = relationship.Active ?? false
                        };
                    }
                }));

                if (foundRelationship == null)
                {
                    throw new InvalidOperationException(
                        $"Relationship not found: {fromTable}.{fromColumn} -> {toTable}.{toColumn}");
                }

                result.Relationship = foundRelationship;
                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelTableColumnsResult ListColumns(IExcelBatch batch, string tableName)
    {
        var result = new DataModelTableColumnsResult
        {
            FilePath = batch.WorkbookPath,
            TableName = tableName
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? table = null;
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

                // Iterate through columns
                ComUtilities.ForEachColumn(table, (Action<dynamic, int>)((column, index) =>
                {
                    bool isCalculated = false;
                    try
                    {
                        // IsCalculatedColumn property may not exist in older Excel versions
                        isCalculated = column.IsCalculatedColumn ?? false;
                    }
                    catch (Exception ex) when (ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException
                                            or System.Runtime.InteropServices.COMException)
                    {
                        // Ignore - property not available in this Excel version
                        isCalculated = false;
                    }

                    var columnInfo = new DataModelColumnInfo
                    {
                        Name = ComUtilities.SafeGetString(column, "Name"),
                        DataType = ComUtilities.SafeGetString(column, "DataType"),
                        IsCalculated = isCalculated
                    };

                    result.Columns.Add(columnInfo);
                }));

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelTableViewResult ReadTable(IExcelBatch batch, string tableName)
    {
        var result = new DataModelTableViewResult
        {
            FilePath = batch.WorkbookPath,
            TableName = tableName
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            dynamic? table = null;
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

                // Get table properties
                result.SourceName = ComUtilities.SafeGetString(table, "SourceName");
                result.RecordCount = ComUtilities.SafeGetInt(table, "RecordCount");

                // Get columns
                ComUtilities.ForEachColumn(table, (Action<dynamic, int>)((column, index) =>
                {
                    bool isCalculated = false;
                    try
                    {
                        // IsCalculatedColumn property may not exist in older Excel versions
                        isCalculated = column.IsCalculatedColumn ?? false;
                    }
                    catch (Exception ex) when (ex is Microsoft.CSharp.RuntimeBinder.RuntimeBinderException
                                            or System.Runtime.InteropServices.COMException)
                    {
                        // Ignore - property not available in this Excel version
                        isCalculated = false;
                    }

                    var columnInfo = new DataModelColumnInfo
                    {
                        Name = ComUtilities.SafeGetString(column, "Name"),
                        DataType = ComUtilities.SafeGetString(column, "DataType"),
                        IsCalculated = isCalculated
                    };

                    result.Columns.Add(columnInfo);
                }));

                // Count measures in this table
                result.MeasureCount = 0;
                ForEachMeasure(model, (Action<dynamic, int>)((measure, index) =>
                {
                    string measureTableName = ComUtilities.SafeGetString(measure.AssociatedTable, "Name");
                    if (string.Equals(measureTableName, tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        result.MeasureCount++;
                    }
                }));

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref table);
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    /// <inheritdoc />
    public DataModelInfoResult ReadInfo(IExcelBatch batch)
    {
        var result = new DataModelInfoResult { FilePath = batch.WorkbookPath };

        using var timeoutCts = new CancellationTokenSource(TimeSpan.FromMinutes(2));

        return batch.Execute((ctx, ct) =>
        {
            dynamic? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                // Count tables and sum rows
                int totalRows = 0;
                ForEachTable(model, (Action<dynamic, int>)((table, index) =>
                {
                    result.TableCount++;
                    totalRows += ComUtilities.SafeGetInt(table, "RecordCount");
                    result.TableNames.Add(ComUtilities.SafeGetString(table, "Name"));
                }));
                result.TotalRows = totalRows;

                // Count measures
                ForEachMeasure(model, (Action<dynamic, int>)((measure, index) =>
                {
                    result.MeasureCount++;
                }));

                // Count relationships
                ForEachRelationship(model, (Action<dynamic, int>)((relationship, index) =>
                {
                    result.RelationshipCount++;
                }));

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        }, timeoutCts.Token);
    }
}

