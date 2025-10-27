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
}
