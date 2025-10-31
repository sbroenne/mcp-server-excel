using Sbroenne.ExcelMcp.ComInterop;
using System;
using System.Collections.Generic;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Private helper methods for DataModel commands
/// </summary>
public partial class DataModelCommands
{
    /// <summary>
    /// Gets all measure names from the Data Model
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <returns>List of measure names</returns>
    private static List<string> GetModelMeasureNames(dynamic model)
    {
        var names = new List<string>();
        dynamic? measures = null;
        try
        {
            // Get measures collection from MODEL (not from tables!)
            // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.model.modelmeasures
            measures = model.ModelMeasures;

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
        }
        return names;
    }

    /// <summary>
    /// Gets the table name that contains a specific measure
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <param name="measureName">Measure name to find</param>
    /// <returns>Table name if found, null otherwise</returns>
    private static string? GetMeasureTableName(dynamic model, string measureName)
    {
        dynamic? measures = null;
        try
        {
            // Get measures collection from MODEL (not from table!)
            // All measures are at model level with AssociatedTable property
            // Reference: https://learn.microsoft.com/en-us/office/vba/api/excel.model.modelmeasures
            measures = model.ModelMeasures;

            for (int m = 1; m <= measures.Count; m++)
            {
                dynamic? measure = null;
                try
                {
                    measure = measures.Item(m);
                    string name = measure.Name?.ToString() ?? "";
                    if (name.Equals(measureName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Get the associated table name
                        dynamic? associatedTable = null;
                        try
                        {
                            associatedTable = measure.AssociatedTable;
                            return associatedTable?.Name?.ToString() ?? string.Empty;
                        }
                        finally
                        {
                            ComUtilities.Release(ref associatedTable);
                        }
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
        }
        return null;
    }

    /// <summary>
    /// Finds a column in a model table by name (case-insensitive)
    /// </summary>
    /// <param name="table">Table COM object</param>
    /// <param name="columnName">Column name to find</param>
    /// <returns>Column COM object or null if not found</returns>
    private static dynamic? FindModelTableColumn(dynamic table, string columnName)
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
                        // Found match - don't release, caller will use it
                        var foundColumn = column;
                        column = null;  // Prevent release in finally
                        return foundColumn;
                    }
                }
                finally
                {
                    // Only release if we didn't return it
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
    private static dynamic? GetFormatObject(dynamic model, string? formatType)
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
    /// Finds a relationship between two tables by column names
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <param name="fromTable">From table name</param>
    /// <param name="fromColumn">From column name</param>
    /// <param name="toTable">To table name</param>
    /// <param name="toColumn">To column name</param>
    /// <returns>Relationship COM object or null if not found</returns>
    private static dynamic? FindRelationship(dynamic model, string fromTable, string fromColumn, string toTable, string toColumn)
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
                        // Found match - don't release, caller will use it
                        var foundRelationship = relationship;
                        relationship = null;  // Prevent release in finally
                        return foundRelationship;
                    }
                }
                finally
                {
                    // Only release if we didn't return it
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

    // ==================== DATA MODEL COM OPERATIONS ====================


    /// <summary>
    /// Checks if the Data Model has any tables
    /// NOTE: Every workbook has a Model object, but it may be empty
    /// </summary>
    private static bool HasDataModelTables(dynamic workbook)
    {
        dynamic? model = null;
        dynamic? modelTables = null;
        try
        {
            model = workbook.Model;
            modelTables = model.ModelTables;
            return modelTables != null && modelTables.Count > 0;
        }
        catch
        {
            return false;
        }
        finally
        {
            ComUtilities.Release(ref modelTables);
            ComUtilities.Release(ref model);
        }
    }

    /// <summary>
    /// Finds a table in the Data Model by name
    /// </summary>
    private static dynamic? FindModelTable(dynamic model, string tableName)
    {
        dynamic? modelTables = null;
        try
        {
            modelTables = model.ModelTables;
            for (int i = 1; i <= modelTables.Count; i++)
            {
                dynamic? table = null;
                try
                {
                    table = modelTables.Item(i);
                    string name = table.Name?.ToString() ?? "";
                    if (name.Equals(tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        var result = table;
                        table = null; // Don't release - returning it
                        return result;
                    }
                }
                finally
                {
                    if (table != null) ComUtilities.Release(ref table);
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
    /// Finds a DAX measure by name across all tables in the model
    /// </summary>
    private static dynamic? FindModelMeasure(dynamic model, string measureName)
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

                    try
                    {
                        measures = table.ModelMeasures;
                    }
                    catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
                    {
                        // ModelMeasures property not available on this table (empty Data Model)
                        continue;
                    }
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // ModelMeasures collection not initialized
                        continue;
                    }

                    for (int m = 1; m <= measures.Count; m++)
                    {
                        dynamic? measure = null;
                        try
                        {
                            measure = measures.Item(m);
                            string name = measure.Name?.ToString() ?? "";
                            if (name.Equals(measureName, StringComparison.OrdinalIgnoreCase))
                            {
                                var result = measure;
                                measure = null; // Don't release - returning it
                                return result;
                            }
                        }
                        finally
                        {
                            if (measure != null) ComUtilities.Release(ref measure);
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
    /// Safely iterates through all tables in the Data Model
    /// </summary>
    private static void ForEachTable(dynamic model, Action<dynamic, int> action)
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
    /// Safely iterates through all measures in the Data Model
    /// </summary>
    private static void ForEachMeasure(dynamic model, Action<dynamic, int> action)
    {
        dynamic? measures = null;
        try
        {
            // Get measures collection from MODEL (not from table!)
            measures = model.ModelMeasures;
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
    /// Safely iterates through all relationships in the Data Model
    /// </summary>
    private static void ForEachRelationship(dynamic model, Action<dynamic, int> action)
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
}
