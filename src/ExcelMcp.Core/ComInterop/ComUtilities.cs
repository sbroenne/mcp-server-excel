using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core.ComInterop;

/// <summary>
/// Low-level COM interop utilities for Excel automation.
/// Provides helpers for finding Excel objects and managing COM object lifecycle.
/// </summary>
public static class ComUtilities
{
    /// <summary>
    /// Safely releases a COM object and sets the reference to null
    /// </summary>
    /// <param name="comObject">The COM object to release</param>
    /// <remarks>
    /// Use this helper to release intermediate COM objects (like ranges, worksheets, queries)
    /// to prevent Excel process from staying open. This is especially important when
    /// iterating through collections or accessing multiple COM properties.
    /// </remarks>
    /// <example>
    /// <code>
    /// dynamic? queries = null;
    /// try
    /// {
    ///     queries = workbook.Queries;
    ///     // Use queries...
    /// }
    /// finally
    /// {
    ///     ComUtilities.Release(ref queries);
    /// }
    /// </code>
    /// </example>
    public static void Release<T>(ref T? comObject) where T : class
    {
        if (comObject != null)
        {
            try
            {
                Marshal.ReleaseComObject(comObject);
            }
            catch
            {
                // Ignore errors during release
            }
            comObject = null;
        }
    }

    /// <summary>
    /// Finds a Power Query by name in the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="queryName">Name of the query to find</param>
    /// <returns>The query COM object if found, null otherwise</returns>
    public static dynamic? FindQuery(dynamic workbook, string queryName)
    {
        dynamic? queriesCollection = null;
        try
        {
            queriesCollection = workbook.Queries;
            int count = queriesCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? query = null;
                try
                {
                    query = queriesCollection.Item(i);
                    if (query.Name == queryName)
                    {
                        // Return the query but don't release it - caller owns it
                        return query;
                    }
                }
                finally
                {
                    // Only release if not returning
                    if (query != null && query.Name != queryName)
                    {
                        Release(ref query);
                    }
                }
            }
        }
        catch { }
        finally
        {
            Release(ref queriesCollection);
        }
        return null;
    }

    /// <summary>
    /// Finds a named range by name in the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="name">Name of the named range to find</param>
    /// <returns>The named range COM object if found, null otherwise</returns>
    public static dynamic? FindName(dynamic workbook, string name)
    {
        dynamic? namesCollection = null;
        try
        {
            namesCollection = workbook.Names;
            int count = namesCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? nameObj = null;
                try
                {
                    nameObj = namesCollection.Item(i);
                    if (nameObj.Name == name)
                    {
                        return nameObj; // Caller owns this object
                    }
                }
                finally
                {
                    // Only release non-matching names
                    if (nameObj != null && nameObj.Name != name)
                    {
                        Release(ref nameObj);
                    }
                }
            }
        }
        catch { }
        finally
        {
            Release(ref namesCollection);
        }
        return null;
    }

    /// <summary>
    /// Finds a worksheet by name in the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="sheetName">Name of the worksheet to find</param>
    /// <returns>The worksheet COM object if found, null otherwise</returns>
    public static dynamic? FindSheet(dynamic workbook, string sheetName)
    {
        dynamic? sheetsCollection = null;
        try
        {
            sheetsCollection = workbook.Worksheets;
            int count = sheetsCollection.Count;
            for (int i = 1; i <= count; i++)
            {
                dynamic? sheet = null;
                try
                {
                    sheet = sheetsCollection.Item(i);
                    if (sheet.Name == sheetName)
                    {
                        return sheet; // Caller owns this object
                    }
                }
                finally
                {
                    // Only release non-matching sheets
                    if (sheet != null && sheet.Name != sheetName)
                    {
                        Release(ref sheet);
                    }
                }
            }
        }
        catch { }
        finally
        {
            Release(ref sheetsCollection);
        }
        return null;
    }

    /// <summary>
    /// Finds a connection in the workbook by name (case-insensitive)
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="connectionName">Name of the connection to find</param>
    /// <returns>Connection object if found, null otherwise</returns>
    public static dynamic? FindConnection(dynamic workbook, string connectionName)
    {
        try
        {
            dynamic connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic conn = connections.Item(i);
                string name = conn.Name?.ToString() ?? "";

                // Match exact name or "Query - Name" pattern (Power Query connections)
                if (name.Equals(connectionName, StringComparison.OrdinalIgnoreCase) ||
                    name.Equals($"Query - {connectionName}", StringComparison.OrdinalIgnoreCase))
                {
                    return conn;
                }
            }
        }
        catch
        {
            // Return null if any error occurs
        }

        return null;
    }

    /// <summary>
    /// Finds a Data Model table by name
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <param name="tableName">Table name to find</param>
    /// <returns>Table COM object if found, null otherwise</returns>
    public static dynamic? FindModelTable(dynamic model, string tableName)
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
                    if (table != null) Release(ref table);
                }
            }
        }
        finally
        {
            Release(ref modelTables);
        }
        return null;
    }

    /// <summary>
    /// Finds a DAX measure by name across all tables in the model
    /// </summary>
    /// <param name="model">Model COM object</param>
    /// <param name="measureName">Measure name to find</param>
    /// <returns>Measure COM object if found, null otherwise</returns>
    public static dynamic? FindModelMeasure(dynamic model, string measureName)
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
                                var result = measure;
                                measure = null; // Don't release - returning it
                                return result;
                            }
                        }
                        finally
                        {
                            if (measure != null) Release(ref measure);
                        }
                    }
                }
                finally
                {
                    Release(ref measures);
                    Release(ref table);
                }
            }
        }
        finally
        {
            Release(ref modelTables);
        }
        return null;
    }
}
