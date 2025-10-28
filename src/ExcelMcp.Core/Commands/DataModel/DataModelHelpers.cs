using System;
using System.Collections.Generic;
using Sbroenne.ExcelMcp.ComInterop;

namespace Sbroenne.ExcelMcp.Core.DataModel;

/// <summary>
/// Helper methods for Excel Data Model operations
/// </summary>
public static class DataModelHelpers
{
    /// <summary>
    /// Checks if a workbook has a Data Model
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <returns>True if Data Model exists</returns>
    public static bool HasDataModel(dynamic workbook)
    {
        dynamic? model = null;
        try
        {
            model = workbook.Model;
            if (model == null) return false;

            // Try to access model tables to confirm model is accessible
            dynamic? modelTables = null;
            try
            {
                modelTables = model.ModelTables;
                return modelTables != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                ComUtilities.Release(ref modelTables);
            }
        }
        catch
        {
            return false;
        }
        finally
        {
            ComUtilities.Release(ref model);
        }
    }

    /// <summary>
    /// Gets all measure names from the Data Model
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <returns>List of measure names</returns>
    public static List<string> GetModelMeasureNames(dynamic model)
    {
        var names = new List<string>();
        dynamic? modelTables = null;
        try
        {
            modelTables = model.ModelTables;
            for (int t = 1; t <= modelTables.Count; t++)
            {
                dynamic? table = null;
                dynamic? measures = null;
                try
                {
                    table = modelTables.Item(t);
                    measures = table.ModelMeasures;

                    for (int m = 1; m <= measures.Count; m++)
                    {
                        dynamic? measure = null;
                        try
                        {
                            measure = measures.Item(m);
                            names.Add(measure.Name?.ToString() ?? "");
                        }
                        finally
                        {
                            ComUtilities.Release(ref measure);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref measures);
                    ComUtilities.Release(ref table);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
        }
        return names;
    }

    /// <summary>
    /// Gets the table name that contains a specific measure
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <param name="measureName">Measure name to find</param>
    /// <returns>Table name if found, null otherwise</returns>
    public static string? GetMeasureTableName(dynamic model, string measureName)
    {
        dynamic? modelTables = null;
        try
        {
            modelTables = model.ModelTables;
            for (int t = 1; t <= modelTables.Count; t++)
            {
                dynamic? table = null;
                dynamic? measures = null;
                try
                {
                    table = modelTables.Item(t);
                    string tableName = table.Name?.ToString() ?? "";
                    measures = table.ModelMeasures;

                    for (int m = 1; m <= measures.Count; m++)
                    {
                        dynamic? measure = null;
                        try
                        {
                            measure = measures.Item(m);
                            string name = measure.Name?.ToString() ?? "";
                            if (name.Equals(measureName, StringComparison.OrdinalIgnoreCase))
                            {
                                return tableName;
                            }
                        }
                        finally
                        {
                            ComUtilities.Release(ref measure);
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref measures);
                    ComUtilities.Release(ref table);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
        }
        return null;
    }

    /// <summary>
    /// Safely iterates through all tables in the Data Model with automatic COM cleanup
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <param name="action">Action to perform on each table</param>
    public static void ForEachTable(dynamic model, Action<dynamic, int> action)
    {
        dynamic? modelTables = null;
        try
        {
            modelTables = model.ModelTables;
            int count = modelTables.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic? table = null;
                try
                {
                    table = modelTables.Item(i);
                    action(table, i);
                }
                finally
                {
                    ComUtilities.Release(ref table);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
        }
    }

    /// <summary>
    /// Safely iterates through all measures in a table with automatic COM cleanup
    /// </summary>
    /// <param name="table">Table COM object</param>
    /// <param name="action">Action to perform on each measure</param>
    public static void ForEachMeasure(dynamic table, Action<dynamic, int> action)
    {
        dynamic? measures = null;
        try
        {
            measures = table.ModelMeasures;
            int count = measures.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic? measure = null;
                try
                {
                    measure = measures.Item(i);
                    action(measure, i);
                }
                finally
                {
                    ComUtilities.Release(ref measure);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref measures);
        }
    }

    /// <summary>
    /// Safely iterates through all relationships in the Data Model with automatic COM cleanup
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <param name="action">Action to perform on each relationship</param>
    public static void ForEachRelationship(dynamic model, Action<dynamic, int> action)
    {
        dynamic? relationships = null;
        try
        {
            relationships = model.ModelRelationships;
            int count = relationships.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic? relationship = null;
                try
                {
                    relationship = relationships.Item(i);
                    action(relationship, i);
                }
                finally
                {
                    ComUtilities.Release(ref relationship);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref relationships);
        }
    }

    /// <summary>
    /// Safely gets a string property from a COM object, returning empty string if null
    /// </summary>
    /// <param name="obj">COM object</param>
    /// <param name="propertyName">Property name</param>
    /// <returns>Property value or empty string</returns>
    public static string SafeGetString(dynamic obj, string propertyName)
    {
        try
        {
            var value = propertyName switch
            {
                "Name" => obj.Name,
                "Formula" => obj.Formula,
                "Description" => obj.Description,
                "SourceName" => obj.SourceName,
                _ => null
            };
            return value?.ToString() ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    /// <summary>
    /// Safely gets an integer property from a COM object, returning 0 if null or invalid
    /// </summary>
    /// <param name="obj">COM object</param>
    /// <param name="propertyName">Property name</param>
    /// <returns>Property value or 0</returns>
    public static int SafeGetInt(dynamic obj, string propertyName)
    {
        try
        {
            var value = propertyName switch
            {
                "RecordCount" => obj.RecordCount,
                "Count" => obj.Count,
                _ => 0
            };
            return Convert.ToInt32(value);
        }
        catch
        {
            return 0;
        }
    }

    /// <summary>
    /// Finds a column in a model table by name (case-insensitive)
    /// </summary>
    /// <param name="table">Table COM object</param>
    /// <param name="columnName">Column name to find</param>
    /// <returns>Column COM object or null if not found</returns>
    public static dynamic? FindModelTableColumn(dynamic table, string columnName)
    {
        dynamic? columns = null;
        try
        {
            columns = table.ModelTableColumns;
            int count = columns.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic? column = null;
                try
                {
                    column = columns.Item(i);
                    string currentName = column.Name?.ToString() ?? "";

                    if (currentName.Equals(columnName, StringComparison.OrdinalIgnoreCase))
                    {
                        return column;  // Don't release - caller will use it
                    }
                }
                finally
                {
                    if (column != null)
                    {
                        ComUtilities.Release(ref column);
                    }
                }
            }

            return null;
        }
        finally
        {
            ComUtilities.Release(ref columns);
        }
    }

    /// <summary>
    /// Gets the appropriate format object from the model for measure creation
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <param name="formatType">Format type (Currency, Decimal, Percentage, General)</param>
    /// <returns>FormatInformation COM object or null for General format</returns>
    public static dynamic? GetFormatObject(dynamic model, string? formatType)
    {
        if (string.IsNullOrEmpty(formatType) || formatType.Equals("General", StringComparison.OrdinalIgnoreCase))
        {
            return null;  // General format (no format object needed)
        }

        try
        {
            return formatType.ToLowerInvariant() switch
            {
                "currency" => model.ModelFormatCurrency,
                "decimal" => model.ModelFormatDecimalNumber,
                "percentage" => model.ModelFormatPercentageNumber,
                "wholenumber" => model.ModelFormatWholeNumber,
                _ => null
            };
        }
        catch
        {
            return null;  // Format not available in this Excel version
        }
    }

    /// <summary>
    /// Safely iterates through all columns in a model table with automatic COM cleanup
    /// </summary>
    /// <param name="table">Table COM object</param>
    /// <param name="action">Action to perform on each column</param>
    public static void ForEachColumn(dynamic table, Action<dynamic, int> action)
    {
        dynamic? columns = null;
        try
        {
            columns = table.ModelTableColumns;
            int count = columns.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic? column = null;
                try
                {
                    column = columns.Item(i);
                    action(column, i);
                }
                finally
                {
                    ComUtilities.Release(ref column);
                }
            }
        }
        finally
        {
            ComUtilities.Release(ref columns);
        }
    }

    /// <summary>
    /// Finds a relationship in the Data Model by table and column names (case-insensitive)
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <param name="fromTable">From table name</param>
    /// <param name="fromColumn">From column name</param>
    /// <param name="toTable">To table name</param>
    /// <param name="toColumn">To column name</param>
    /// <returns>Relationship COM object or null if not found</returns>
    public static dynamic? FindRelationship(dynamic model, string fromTable, string fromColumn, string toTable, string toColumn)
    {
        dynamic? relationships = null;
        try
        {
            relationships = model.ModelRelationships;
            int count = relationships.Count;

            for (int i = 1; i <= count; i++)
            {
                dynamic? relationship = null;
                try
                {
                    relationship = relationships.Item(i);
                    
                    // Get relationship details
                    string currentFromTable = relationship.ForeignKeyColumn?.Parent?.Name?.ToString() ?? "";
                    string currentFromColumn = relationship.ForeignKeyColumn?.Name?.ToString() ?? "";
                    string currentToTable = relationship.PrimaryKeyColumn?.Parent?.Name?.ToString() ?? "";
                    string currentToColumn = relationship.PrimaryKeyColumn?.Name?.ToString() ?? "";

                    // Match all four components (case-insensitive)
                    if (currentFromTable.Equals(fromTable, StringComparison.OrdinalIgnoreCase) &&
                        currentFromColumn.Equals(fromColumn, StringComparison.OrdinalIgnoreCase) &&
                        currentToTable.Equals(toTable, StringComparison.OrdinalIgnoreCase) &&
                        currentToColumn.Equals(toColumn, StringComparison.OrdinalIgnoreCase))
                    {
                        return relationship;  // Don't release - caller will use it
                    }
                }
                finally
                {
                    if (relationship != null)
                    {
                        ComUtilities.Release(ref relationship);
                    }
                }
            }

            return null;
        }
        finally
        {
            ComUtilities.Release(ref relationships);
        }
    }
}
