using System.Globalization;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;
using Excel = Microsoft.Office.Interop.Excel;


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

            Excel.Model? model = null;
            try
            {
                model = ctx.Book.Model;

                ForEachTable(model!, (table, index) =>
                {
                    var tableInfo = new DataModelTableInfo
                    {
                        Name = table.Name ?? string.Empty,
                        SourceName = table.SourceName ?? string.Empty,
                        RecordCount = table.RecordCount
                    };

                    result.Tables.Add(tableInfo);
                });

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
            Excel.Model? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;
                var engineFormulas = TryReadStoredMeasureFormulas(model!);
                var measureTables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Iterate through all measures (they're at model level)
                ForEachMeasure(model!, (measure, index) =>
                {
                    // Get the table name for this measure
                    string measureTableName = string.Empty;
                    Excel.ModelTable? associatedTable = null;
                    try
                    {
                        associatedTable = measure.AssociatedTable;
                        measureTableName = associatedTable?.Name?.ToString() ?? string.Empty;
                    }
                    finally
                    {
                        ComUtilities.Release(ref associatedTable);
                    }

                    if (!string.IsNullOrWhiteSpace(measureTableName))
                        measureTables.Add(measureTableName);

                    // Skip if filtering by table and this measure isn't in that table
                    if (tableName != null && !measureTableName.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    string measureName = measure.Name ?? string.Empty;
                    string formula = engineFormulas != null && engineFormulas.TryGetValue(measureName, out var stored)
                        ? stored
                        : measure.Formula ?? string.Empty;

                    var measureInfo = new DataModelMeasureInfo
                    {
                        Name = measureName,
                        Table = measureTableName,
                        FormulaPreview = formula, // Will be formatted after retrieval
                        Description = measure.Description ?? string.Empty
                    };

                    result.Measures.Add(measureInfo);
                });

                // Check if table filter was specified but not found
                if (tableName != null && result.Measures.Count == 0)
                {
                    Excel.ModelTable? requestedTable = null;
                    try
                    {
                        requestedTable = FindModelTable(model!, tableName);
                        if (requestedTable is null)
                            throw new InvalidOperationException(DataModelErrorMessages.TableNotFound(tableName));
                    }
                    finally
                    {
                        ComUtilities.Release(ref requestedTable);
                    }

                    var otherTables = measureTables
                        .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
                        .ToArray();
                    result.Message = otherTables.Length == 0
                        ? $"Table '{tableName}' has no measures. The Data Model contains no measures."
                        : $"Table '{tableName}' has no measures. Measures are stored on: {string.Join(", ", otherTables)}.";
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

                // Get measure details using safe helpers
                string storedName = measure.Name ?? string.Empty;
                var engineFormulas = TryReadStoredMeasureFormulas(model!);
                result.DaxFormula = engineFormulas != null && engineFormulas.TryGetValue(storedName, out var stored)
                    ? stored
                    : measure.Formula ?? string.Empty;
                result.Description = measure.Description ?? string.Empty;
                result.CharacterCount = result.DaxFormula.Length;
                result.TableName = GetMeasureTableName(model!, measureName) ?? "";

                object? formatInfo = null;
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
            Excel.Model? model = null;
            try
            {
                // Check if workbook has Data Model
                if (!HasDataModelTables(ctx.Book))
                {
                    throw new InvalidOperationException(DataModelErrorMessages.NoDataModelTables());
                }

                model = ctx.Book.Model;

                ForEachRelationship(model!, (relationship, index) =>
                {
                    var relInfo = ReadRelationshipInfo(relationship);

                    result.Relationships.Add(relInfo);
                });

                result.Success = true;
            }
            finally
            {
                ComUtilities.Release(ref model);
            }

            return result;
        });
    }

    private static DataModelRelationshipInfo ReadRelationshipInfo(Excel.ModelRelationship relationship)
    {
        Excel.ModelTableColumn? foreignKey = null;
        Excel.ModelTableColumn? primaryKey = null;
        Excel.ModelTable? fromTable = null;
        Excel.ModelTable? toTable = null;
        try
        {
            foreignKey = relationship.ForeignKeyColumn;
            RelationshipReferenceAcquiredForTests?.Invoke(foreignKey);
            primaryKey = relationship.PrimaryKeyColumn;
            RelationshipReferenceAcquiredForTests?.Invoke(primaryKey);
            fromTable = (Excel.ModelTable)foreignKey.Parent;
            RelationshipReferenceAcquiredForTests?.Invoke(fromTable);
            toTable = (Excel.ModelTable)primaryKey.Parent;
            RelationshipReferenceAcquiredForTests?.Invoke(toTable);
            return new DataModelRelationshipInfo
            {
                FromTable = fromTable.Name,
                FromColumn = foreignKey.Name,
                ToTable = toTable.Name,
                ToColumn = primaryKey.Name,
                IsActive = relationship.Active
            };
        }
        finally
        {
            ComUtilities.Release(ref toTable);
            ComUtilities.Release(ref fromTable);
            ComUtilities.Release(ref primaryKey);
            ComUtilities.Release(ref foreignKey);
        }
    }

    /// <inheritdoc />
    public DataModelRelationshipViewResult ReadRelationship(IExcelBatch batch, string fromTable, string fromColumn, string toTable, string toColumn)
    {
        var result = new DataModelRelationshipViewResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            Excel.Model? model = null;
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
                ForEachRelationship(model!, (relationship, index) =>
                {
                    var info = ReadRelationshipInfo(relationship);

                    if (string.Equals(info.FromTable, fromTable, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(info.FromColumn, fromColumn, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(info.ToTable, toTable, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(info.ToColumn, toColumn, StringComparison.OrdinalIgnoreCase))
                    {
                        foundRelationship = info;
                    }
                });

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
            Excel.Model? model = null;
            Excel.ModelTable? table = null;
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

                // Iterate through columns
                ComUtilities.ForEachColumn(table, (column, index) =>
                {
                    result.Columns.Add(CreateColumnInfo(column));
                });

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
            Excel.Model? model = null;
            Excel.ModelTable? table = null;
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

                // Get table properties
                result.SourceName = table.SourceName ?? string.Empty;
                result.RecordCount = table.RecordCount;

                Excel.WorkbookConnection? sourceConnection = null;
                try
                {
                    sourceConnection = table.SourceWorkbookConnection;
                    result.SourceConnectionName = sourceConnection.Name ?? string.Empty;
                    result.SourceConnectionDescription = sourceConnection.Description ?? string.Empty;
                    result.SourceConnectionTypeValue = Convert.ToInt32(sourceConnection.Type, CultureInfo.InvariantCulture);
                    result.SourceConnectionType = ConnectionHelpers.GetConnectionTypeName(result.SourceConnectionTypeValue);
                    result.SourceConnectionInModel = sourceConnection.InModel;
                }
                finally
                {
                    ComUtilities.Release(ref sourceConnection);
                }

                // Get columns
                ComUtilities.ForEachColumn(table, (column, index) =>
                {
                    result.Columns.Add(CreateColumnInfo(column));
                });

                // Count measures in this table
                result.MeasureCount = 0;
                ForEachMeasure(model!, (measure, index) =>
                {
                    Excel.ModelTable? associatedTable = null;
                    try
                    {
                        associatedTable = measure.AssociatedTable;
                        string measureTableName = associatedTable.Name ?? string.Empty;
                        if (string.Equals(measureTableName, tableName, StringComparison.OrdinalIgnoreCase))
                        {
                            result.MeasureCount++;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref associatedTable);
                    }
                });

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
            Excel.Model? model = null;
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
                ForEachTable(model!, (table, index) =>
                {
                    result.TableCount++;
                    totalRows += table.RecordCount;
                    result.TableNames.Add(table.Name ?? string.Empty);
                });
                result.TotalRows = totalRows;

                // Count measures
                ForEachMeasure(model!, (measure, index) =>
                {
                    result.MeasureCount++;
                });

                // Count relationships
                ForEachRelationship(model!, (relationship, index) =>
                {
                    result.RelationshipCount++;
                });

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
