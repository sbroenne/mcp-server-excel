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
            try
            {
                measures = model.ModelMeasures;
            }
            catch (Microsoft.CSharp.RuntimeBinder.RuntimeBinderException)
            {
                // ModelMeasures API not available (requires Office 2016+)
                throw new InvalidOperationException(
                    "DAX measures are not supported in this version of Excel. " +
                    "The ModelMeasures API requires Microsoft Office 2016 or later. " +
                    "Please upgrade Excel to use measure operations.");
            }

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
}
