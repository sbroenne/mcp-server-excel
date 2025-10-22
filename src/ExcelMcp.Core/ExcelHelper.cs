using System.Diagnostics.CodeAnalysis;
using System.Runtime.InteropServices;

namespace Sbroenne.ExcelMcp.Core;

/// <summary>
/// Helper class for Excel COM automation with proper resource management
/// </summary>
public static class ExcelHelper
{
    /// <summary>
    /// Optional Excel instance pool for improved performance in conversational workflows.
    /// When set, WithExcel will use pooled instances instead of creating new Excel instances.
    /// This is automatically configured by MCP Server for optimal AI assistant performance.
    /// </summary>
    public static ExcelInstancePool? InstancePool { get; set; }

    /// <summary>
    /// Executes an action with Excel COM automation using proper resource management.
    /// Automatically uses pooled instances if InstancePool is configured, providing
    /// significant performance improvements for conversational workflows (~2-5 second
    /// startup overhead reduced to near-instantaneous for cached workbooks).
    /// </summary>
    /// <typeparam name="T">Return type of the action</typeparam>
    /// <param name="filePath">Path to the Excel file</param>
    /// <param name="save">Whether to save changes to the file</param>
    /// <param name="action">Action to execute with Excel application and workbook</param>
    /// <returns>Result of the action</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static T WithExcel<T>(string filePath, bool save, Func<dynamic, dynamic, T> action)
    {
        // Use pooled instance if available (MCP Server optimization)
        var pool = InstancePool;
        if (pool != null)
        {
            return pool.WithPooledExcel(filePath, save, action);
        }

        // Fall back to single-instance pattern (CLI and backward compatibility)
        return WithExcelSingleInstance(filePath, save, action);
    }

    /// <summary>
    /// Single-instance Excel execution pattern - creates new Excel instance for each operation.
    /// This is the traditional pattern used by CLI commands for simplicity and reliability.
    /// </summary>
    private static T WithExcelSingleInstance<T>(string filePath, bool save, Func<dynamic, dynamic, T> action)
    {
        dynamic? excel = null;
        dynamic? workbook = null;
        string operation = $"WithExcel({Path.GetFileName(filePath)}, save={save})";

        try
        {
            // Validate file path first - prevent path traversal attacks
            string fullPath = Path.GetFullPath(filePath);

            // Additional security: ensure the file is within reasonable bounds
            if (fullPath.Length > 32767)
            {
                throw new ArgumentException($"File path too long: {fullPath.Length} characters (Windows limit: 32767)");
            }

            // Security: Validate file extension to prevent executing arbitrary files
            string extension = Path.GetExtension(fullPath).ToLowerInvariant();
            if (extension is not (".xlsx" or ".xlsm" or ".xls"))
            {
                throw new ArgumentException($"Invalid file extension '{extension}'. Only Excel files (.xlsx, .xlsm, .xls) are supported.");
            }

            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"Excel file not found: {fullPath}", fullPath);
            }

            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException("Excel is not installed or not properly registered. " +
                    "Please verify Microsoft Excel is installed and COM registration is intact.");
            }

#pragma warning disable IL2072 // COM interop is not AOT compatible but is required for Excel automation
            excel = Activator.CreateInstance(excelType);
#pragma warning restore IL2072
            if (excel == null)
            {
                throw new InvalidOperationException("Failed to create Excel COM instance. " +
                    "Excel may be corrupted or COM subsystem unavailable.");
            }

            // Configure Excel for automation
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.Interactive = false;

            // Open workbook with detailed error context
            try
            {
                workbook = excel.Workbooks.Open(fullPath);
            }
            catch (COMException comEx) when (comEx.ErrorCode == unchecked((int)0x8001010A))
            {
                // Excel is busy - provide specific guidance
                throw new InvalidOperationException(
                    "Excel is busy (likely has a dialog open). Close any Excel dialogs and retry.", comEx);
            }
            catch (COMException comEx) when (comEx.ErrorCode == unchecked((int)0x80070020))
            {
                // File sharing violation
                throw new InvalidOperationException(
                    $"File '{Path.GetFileName(fullPath)}' is locked by another process. " +
                    "Close Excel and any other applications using this file.", comEx);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"Failed to open workbook '{Path.GetFileName(fullPath)}'. " +
                    "File may be corrupted, password-protected, or incompatible.", ex);
            }

            if (workbook == null)
            {
                throw new InvalidOperationException($"Failed to open workbook: {Path.GetFileName(fullPath)}");
            }

            // Execute the user action with error context and retry logic for transient errors
            T result;
            int retryCount = 0;
            const int maxRetries = 3;
            
            while (true)
            {
                try
                {
                    result = action(excel, workbook);
                    break; // Success - exit retry loop
                }
                catch (COMException comEx) when (comEx.HResult == unchecked((int)0x8001010A) && retryCount < maxRetries)
                {
                    // Excel is busy (RPC_E_SERVERCALL_RETRYLATER)
                    // This can happen during parallel operations or when Excel is processing
                    retryCount++;
                    System.Threading.Thread.Sleep(500 * retryCount); // Exponential backoff: 500ms, 1s, 1.5s
                    
                    if (retryCount >= maxRetries)
                    {
                        throw new InvalidOperationException(
                            "Excel is busy. Please close any dialogs and try again.", comEx);
                    }
                    // Continue retry loop
                }
                catch
                {
                    // Propagate all other exceptions with original context
                    throw;
                }
            }

            // Save if requested
            if (save && workbook != null)
            {
                try
                {
                    workbook.Save();
                }
                catch
                {
                    // Propagate save exceptions
                    throw;
                }
            }

            return result;
        }
        catch
        {
            // Propagate exceptions to caller
            throw;
        }
        finally
        {
            // Enhanced COM cleanup to prevent process leaks

            // Close workbook first
            if (workbook != null)
            {
                try
                {
                    workbook.Close(save);
                }
                catch (COMException)
                {
                    // Workbook might already be closed, ignore
                }
                catch
                {
                    // Any other exception during close, ignore to continue cleanup
                }

                try
                {
                    Marshal.ReleaseComObject(workbook);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }

                workbook = null;
            }

            // Quit Excel application
            if (excel != null)
            {
                try
                {
                    excel.Quit();
                }
                catch (COMException)
                {
                    // Excel might already be closing, ignore
                }
                catch
                {
                    // Any other exception during quit, ignore to continue cleanup
                }

                try
                {
                    Marshal.ReleaseComObject(excel);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }

                excel = null;
            }

            // Recommended COM cleanup pattern:
            // Two GC cycles are sufficient - one to collect, one to finalize
            // Microsoft recommends against excessive GC.Collect() calls
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect(); // Final collect to clean up objects queued during finalization
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
                        ReleaseComObject(ref query);
                    }
                }
            }
        }
        catch { }
        finally
        {
            ReleaseComObject(ref queriesCollection);
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
                        ReleaseComObject(ref nameObj);
                    }
                }
            }
        }
        catch { }
        finally
        {
            ReleaseComObject(ref namesCollection);
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
                        ReleaseComObject(ref sheet);
                    }
                }
            }
        }
        catch { }
        finally
        {
            ReleaseComObject(ref sheetsCollection);
        }
        return null;
    }

    /// <summary>
    /// Creates a new Excel workbook with proper resource management
    /// </summary>
    /// <typeparam name="T">Return type of the action</typeparam>
    /// <param name="filePath">Path where to save the new Excel file</param>
    /// <param name="isMacroEnabled">Whether to create a macro-enabled workbook (.xlsm)</param>
    /// <param name="action">Action to execute with Excel application and new workbook</param>
    /// <returns>Result of the action</returns>
    [SuppressMessage("Interoperability", "CA1416:Validate platform compatibility")]
    public static T WithNewExcel<T>(string filePath, bool isMacroEnabled, Func<dynamic, dynamic, T> action)
    {
        dynamic? excel = null;
        dynamic? workbook = null;
        string operation = $"WithNewExcel({Path.GetFileName(filePath)}, macroEnabled={isMacroEnabled})";

        try
        {
            // Validate file path first - prevent path traversal attacks
            string fullPath = Path.GetFullPath(filePath);

            // Validate file size limits for security (prevent DoS)
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Get Excel COM type
            var excelType = Type.GetTypeFromProgID("Excel.Application");
            if (excelType == null)
            {
                throw new InvalidOperationException("Excel is not installed or not properly registered. " +
                    "Please verify Microsoft Excel is installed and COM registration is intact.");
            }

#pragma warning disable IL2072 // COM interop is not AOT compatible but is required for Excel automation
            excel = Activator.CreateInstance(excelType);
#pragma warning restore IL2072
            if (excel == null)
            {
                throw new InvalidOperationException("Failed to create Excel COM instance. " +
                    "Excel may be corrupted or COM subsystem unavailable.");
            }

            // Configure Excel for automation
            excel.Visible = false;
            excel.DisplayAlerts = false;
            excel.ScreenUpdating = false;
            excel.Interactive = false;

            // Create new workbook
            workbook = excel.Workbooks.Add();

            // Execute the user action
            var result = action(excel, workbook);

            // Save the workbook with appropriate format
            if (isMacroEnabled)
            {
                // Save as macro-enabled workbook (format 52)
                workbook.SaveAs(fullPath, 52);
            }
            else
            {
                // Save as regular workbook (format 51)
                workbook.SaveAs(fullPath, 51);
            }

            return result;
        }
        catch (COMException comEx)
        {
            throw new InvalidOperationException($"Excel COM operation failed during {operation}: {comEx.Message}", comEx);
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Operation failed during {operation}: {ex.Message}", ex);
        }
        finally
        {
            // Enhanced COM cleanup to prevent process leaks

            // Close workbook first
            if (workbook != null)
            {
                try
                {
                    workbook.Close(false); // Don't save again, we already saved
                }
                catch (COMException)
                {
                    // Workbook might already be closed, ignore
                }
                catch
                {
                    // Any other exception during close, ignore to continue cleanup
                }

                try
                {
                    Marshal.ReleaseComObject(workbook);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }

                workbook = null;
            }

            // Quit Excel application
            if (excel != null)
            {
                try
                {
                    excel.Quit();
                }
                catch (COMException)
                {
                    // Excel might already be closing, ignore
                }
                catch
                {
                    // Any other exception during quit, ignore to continue cleanup
                }

                try
                {
                    Marshal.ReleaseComObject(excel);
                }
                catch
                {
                    // Release might fail, but continue cleanup
                }

                excel = null;
            }

            // Recommended COM cleanup pattern:
            // Two GC cycles are sufficient - one to collect, one to finalize
            // Microsoft recommends against excessive GC.Collect() calls
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect(); // Final collect to clean up objects queued during finalization
        }
    }

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
    ///     ReleaseComObject(ref queries);
    /// }
    /// </code>
    /// </example>
    public static void ReleaseComObject<T>(ref T? comObject) where T : class
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

    #region Connection and Query Shared Utilities

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
    /// Gets all connection names from the workbook
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <returns>List of connection names</returns>
    public static List<string> GetConnectionNames(dynamic workbook)
    {
        var names = new List<string>();

        try
        {
            dynamic connections = workbook.Connections;
            for (int i = 1; i <= connections.Count; i++)
            {
                dynamic conn = connections.Item(i);
                string name = conn.Name?.ToString() ?? "";
                if (!string.IsNullOrWhiteSpace(name))
                {
                    names.Add(name);
                }
            }
        }
        catch
        {
            // Return empty list if any error occurs
        }

        return names;
    }

    /// <summary>
    /// Gets the connection type name from XlConnectionType enum value
    /// </summary>
    /// <param name="connectionType">Connection type numeric value</param>
    /// <returns>Human-readable connection type name</returns>
    public static string GetConnectionTypeName(int connectionType)
    {
        return connectionType switch
        {
            1 => "OLEDB",
            2 => "ODBC",
            3 => "XML",
            4 => "Text",
            5 => "Web",
            6 => "DataFeed",
            7 => "Model",
            8 => "Worksheet",
            9 => "NoSource",
            _ => $"Unknown ({connectionType})"
        };
    }

    /// <summary>
    /// Determines if a connection is a Power Query connection
    /// </summary>
    /// <param name="connection">Connection COM object</param>
    /// <returns>True if connection is a Power Query connection</returns>
    public static bool IsPowerQueryConnection(dynamic connection)
    {
        try
        {
            // Power Query connections use Microsoft.Mashup provider
            // Check OLEDBConnection for Mashup provider
            if (connection.Type == 1) // xlConnectionTypeOLEDB
            {
                string connectionString = connection.OLEDBConnection?.Connection?.ToString() ?? "";
                if (connectionString.Contains("Microsoft.Mashup.OleDb", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            // Also check connection name pattern (Power Query connections are named "Query - Name")
            string name = connection.Name?.ToString() ?? "";
            if (name.StartsWith("Query - ", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }
        catch
        {
            // If any error occurs, assume not a Power Query connection
        }

        return false;
    }

    /// <summary>
    /// Sanitizes connection string by masking password
    /// SECURITY: Always use this before displaying or exporting connection strings
    /// </summary>
    /// <param name="connectionString">Connection string that may contain password</param>
    /// <returns>Sanitized connection string with password masked</returns>
    public static string SanitizeConnectionString(string? connectionString)
    {
        if (string.IsNullOrWhiteSpace(connectionString))
        {
            return string.Empty;
        }

        // Regex pattern to match password in various formats:
        // Password=value; Pwd=value; password=value; pwd=value;
        // Handles both semicolon-terminated and end-of-string cases
        return System.Text.RegularExpressions.Regex.Replace(
            connectionString,
            @"(password|pwd)\s*=\s*[^;]*",
            "$1=***",
            System.Text.RegularExpressions.RegexOptions.IgnoreCase
        );
    }

    /// <summary>
    /// Removes connections associated with a query or connection name
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="name">Name of the query or connection</param>
    public static void RemoveConnections(dynamic workbook, string name)
    {
        try
        {
            dynamic connections = workbook.Connections;

            // Iterate backwards to safely delete items
            for (int i = connections.Count; i >= 1; i--)
            {
                dynamic conn = connections.Item(i);
                string connName = conn.Name?.ToString() ?? "";

                // Match exact name or "Query - Name" pattern
                if (connName.Equals(name, StringComparison.OrdinalIgnoreCase) ||
                    connName.Equals($"Query - {name}", StringComparison.OrdinalIgnoreCase))
                {
                    conn.Delete();
                }
            }
        }
        catch
        {
            // Ignore errors when removing connections - they may not exist
        }
    }

    /// <summary>
    /// Removes QueryTables associated with a query or connection name from all worksheets
    /// </summary>
    /// <param name="workbook">Excel workbook COM object</param>
    /// <param name="name">Name of the query or connection (spaces will be replaced with underscores for QueryTable names)</param>
    public static void RemoveQueryTables(dynamic workbook, string name)
    {
        try
        {
            dynamic worksheets = workbook.Worksheets;
            string normalizedName = name.Replace(" ", "_");

            for (int ws = 1; ws <= worksheets.Count; ws++)
            {
                dynamic worksheet = worksheets.Item(ws);
                dynamic queryTables = worksheet.QueryTables;

                // Iterate backwards to safely delete items
                for (int qt = queryTables.Count; qt >= 1; qt--)
                {
                    dynamic queryTable = queryTables.Item(qt);
                    string queryTableName = queryTable.Name?.ToString() ?? "";

                    // Match QueryTable names that contain the normalized name
                    if (queryTableName.Contains(normalizedName, StringComparison.OrdinalIgnoreCase))
                    {
                        queryTable.Delete();
                    }
                }
            }
        }
        catch
        {
            // Ignore errors when removing QueryTables - they may not exist
        }
    }

    /// <summary>
    /// Options for creating QueryTable connections
    /// </summary>
    public class QueryTableOptions
    {
        /// <summary>
        /// Name of the query or connection
        /// </summary>
        public required string Name { get; init; }

        /// <summary>
        /// Whether to refresh data in background
        /// </summary>
        public bool BackgroundQuery { get; init; } = false;

        /// <summary>
        /// Whether to refresh data when file opens
        /// </summary>
        public bool RefreshOnFileOpen { get; init; } = false;

        /// <summary>
        /// Whether to save password in connection
        /// </summary>
        public bool SavePassword { get; init; } = false;

        /// <summary>
        /// Whether to preserve column information
        /// </summary>
        public bool PreserveColumnInfo { get; init; } = true;

        /// <summary>
        /// Whether to preserve formatting
        /// </summary>
        public bool PreserveFormatting { get; init; } = true;

        /// <summary>
        /// Whether to auto-adjust column width
        /// </summary>
        public bool AdjustColumnWidth { get; init; } = true;

        /// <summary>
        /// Whether to refresh immediately after creation
        /// </summary>
        public bool RefreshImmediately { get; init; } = false;
    }

    /// <summary>
    /// Creates a QueryTable connection that loads data from a Power Query to a worksheet
    /// </summary>
    /// <param name="targetSheet">Target worksheet COM object</param>
    /// <param name="queryName">Name of the Power Query</param>
    /// <param name="options">QueryTable configuration options</param>
    public static void CreateQueryTable(dynamic targetSheet, string queryName, QueryTableOptions? options = null)
    {
        options ??= new QueryTableOptions { Name = queryName };

        dynamic queryTables = targetSheet.QueryTables;

        // Connection string for Power Query (uses Microsoft.Mashup.OleDb provider)
        string connectionString = $"OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location={queryName}";
        string commandText = $"SELECT * FROM [{queryName}]";

        // Create QueryTable at cell A1
        dynamic queryTable = queryTables.Add(connectionString, targetSheet.Range["A1"], commandText);

        // Configure QueryTable properties
        queryTable.Name = options.Name.Replace(" ", "_");
        queryTable.RefreshStyle = 1; // xlInsertDeleteCells
        queryTable.BackgroundQuery = options.BackgroundQuery;
        queryTable.RefreshOnFileOpen = options.RefreshOnFileOpen;
        queryTable.SavePassword = options.SavePassword;
        queryTable.PreserveColumnInfo = options.PreserveColumnInfo;
        queryTable.PreserveFormatting = options.PreserveFormatting;
        queryTable.AdjustColumnWidth = options.AdjustColumnWidth;

        // Refresh immediately if requested
        if (options.RefreshImmediately)
        {
            queryTable.Refresh(false);
        }
    }

    #endregion

    #region Data Model Helper Methods

    /// <summary>
    /// Checks if workbook has a Data Model
    /// </summary>
    /// <param name="workbook">Workbook COM object</param>
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
                ReleaseComObject(ref modelTables);
            }
        }
        catch
        {
            return false;
        }
        finally
        {
            ReleaseComObject(ref model);
        }
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
                    if (table != null) ReleaseComObject(ref table);
                }
            }
        }
        finally
        {
            ReleaseComObject(ref modelTables);
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
                            if (measure != null) ReleaseComObject(ref measure);
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(ref measures);
                    ReleaseComObject(ref table);
                }
            }
        }
        finally
        {
            ReleaseComObject(ref modelTables);
        }
        return null;
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
                            ReleaseComObject(ref measure);
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(ref measures);
                    ReleaseComObject(ref table);
                }
            }
        }
        finally
        {
            ReleaseComObject(ref modelTables);
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
                            ReleaseComObject(ref measure);
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(ref measures);
                    ReleaseComObject(ref table);
                }
            }
        }
        finally
        {
            ReleaseComObject(ref modelTables);
        }
        return null;
    }

    #endregion

}
