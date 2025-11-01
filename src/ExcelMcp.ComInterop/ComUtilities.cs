using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.ComInterop;

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
    /// <remarks>
    /// CRITICAL: Caller is responsible for releasing the returned COM object.
    /// Use ComUtilities.Release(ref query) when done with the object.
    /// </remarks>
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
                    string currentName = query.Name;

                    if (currentName == queryName)
                    {
                        // Found match - return it (caller owns it now)
                        var result = query;
                        query = null; // Prevent cleanup in finally block
                        return result;
                    }
                }
                finally
                {
                    // Only release if not returning (query will be null if we're returning it)
                    if (query != null)
                    {
                        Release(ref query);
                    }
                }
            }

            return null; // Not found
        }
        catch (Exception ex)
        {
            // Log or rethrow - don't silently swallow
            throw new InvalidOperationException($"Failed to search for Power Query '{queryName}'.", ex);
        }
        finally
        {
            Release(ref queriesCollection);
        }
    }

    /// <summary>
    /// Finds a named range by name in the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="name">Name of the named range to find</param>
    /// <returns>The named range COM object if found, null otherwise</returns>
    /// <remarks>
    /// CRITICAL: Caller is responsible for releasing the returned COM object.
    /// Use ComUtilities.Release(ref nameObj) when done with the object.
    /// </remarks>
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
                    string currentName = nameObj.Name;

                    if (currentName == name)
                    {
                        // Found match - return it (caller owns it now)
                        var result = nameObj;
                        nameObj = null; // Prevent cleanup in finally block
                        return result;
                    }
                }
                finally
                {
                    // Only release if not returning
                    if (nameObj != null)
                    {
                        Release(ref nameObj);
                    }
                }
            }

            return null; // Not found
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to search for named range '{name}'.", ex);
        }
        finally
        {
            Release(ref namesCollection);
        }
    }

    /// <summary>
    /// Finds a worksheet by name in the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="sheetName">Name of the worksheet to find</param>
    /// <returns>The worksheet COM object if found, null otherwise</returns>
    /// <remarks>
    /// CRITICAL: Caller is responsible for releasing the returned COM object.
    /// Use ComUtilities.Release(ref sheet) when done with the object.
    /// </remarks>
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
                    string currentName = sheet.Name;

                    if (currentName == sheetName)
                    {
                        // Found match - return it (caller owns it now)
                        var result = sheet;
                        sheet = null; // Prevent cleanup in finally block
                        return result;
                    }
                }
                finally
                {
                    // Only release if not returning
                    if (sheet != null)
                    {
                        Release(ref sheet);
                    }
                }
            }

            return null; // Not found
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to search for worksheet '{sheetName}'.", ex);
        }
        finally
        {
            Release(ref sheetsCollection);
        }
    }

    /// <summary>
    /// Finds a connection in the workbook by name (case-insensitive)
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="connectionName">Name of the connection to find</param>
    /// <returns>Connection object if found, null otherwise</returns>
    /// <remarks>
    /// CRITICAL: Caller is responsible for releasing the returned COM object.
    /// Use ComUtilities.Release(ref connection) when done with the object.
    /// </remarks>
    public static dynamic? FindConnection(dynamic workbook, string connectionName)
    {
        dynamic? connections = null;
        dynamic? conn = null;

        try
        {
            connections = workbook.Connections;

            for (int i = 1; i <= connections.Count; i++)
            {
                conn = connections.Item(i);
                string name = conn.Name?.ToString() ?? "";

                // Match exact name or "Query - Name" pattern (Power Query connections)
                if (name.Equals(connectionName, StringComparison.OrdinalIgnoreCase) ||
                    name.Equals($"Query - {connectionName}", StringComparison.OrdinalIgnoreCase))
                {
                    // Found match - return it (caller owns it now)
                    var result = conn;
                    conn = null; // Prevent cleanup in finally block
                    return result;
                }

                // Not a match - release before next iteration
                Release(ref conn);
            }

            return null; // Not found
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to search for connection '{connectionName}'.", ex);
        }
        finally
        {
            // Clean up any unreleased connection from last iteration
            if (conn != null)
            {
                Release(ref conn);
            }
            Release(ref connections);
        }
    }

    /// <summary>
    /// Safely iterates through all columns in a model table with automatic COM cleanup
    /// </summary>
    /// <param name="table">Model table COM object</param>
    /// <param name="action">Action to perform on each column (receives column and 1-based index)</param>
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
                    Release(ref column);
                }
            }
        }
        finally
        {
            Release(ref columns);
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
    /// Safely gets a DateTime property from a COM object, returning null if unavailable
    /// </summary>
    /// <param name="obj">COM object</param>
    /// <param name="propertyName">Property name</param>
    /// <returns>Property value or null if not available</returns>
    public static DateTime? SafeGetDateTime(dynamic obj, string propertyName)
    {
        try
        {
            var value = propertyName switch
            {
                "RefreshDate" => obj.RefreshDate,
                _ => null
            };
            
            if (value == null) return null;
            if (value is DateTime dt) return dt;
            if (value is double dbl) return DateTime.FromOADate(dbl);
            
            return null;
        }
        catch
        {
            return null;
        }
    }
}
