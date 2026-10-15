using System.Data.Common;
using System.Globalization;
using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;


namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection management commands - Core data layer (no console output)
/// Provides CRUD operations for Excel data connections (OLEDB, ODBC, Text, Web, etc.)
/// </summary>
public partial class ConnectionCommands : IConnectionCommands
{
    #region Helper Methods

    private static DateTime? GetRefreshDateSafe(object? refreshDate)
    {
        return refreshDate switch
        {
            null => null,
            DateTime dateTime => dateTime,
            double oaDate => DateTime.FromOADate(oaDate),
            _ => null
        };
    }

    private static object? GetTypedSubConnection(Excel.WorkbookConnection conn)
    {
        int connType = (int)conn.Type;

        if (connType == 1)
        {
            return conn.OLEDBConnection;
        }

        if (connType == 2)
        {
            return conn.ODBCConnection;
        }

        if (connType is 3 or 4)
        {
            try
            {
                return conn.TextConnection;
            }
            catch (COMException)
            {
                try
                {
                    // PIA gap: WorkbookConnection.WebConnection is absent from the Excel PIA.

                    return ((dynamic)(object)conn).WebConnection;
                }
                catch (COMException)
                {
                    return null;
                }
            }
        }

        return null;
    }

    private static bool GetBackgroundQuerySetting(Excel.WorkbookConnection conn)
    {
        try
        {
            int connType = (int)conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection != null && !oledbConnection.OLAP && oledbConnection.BackgroundQuery;
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.BackgroundQuery ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                Excel.TextConnection? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    // PIA gap: TextConnection.BackgroundQuery is absent from the Excel PIA.
                    return ((dynamic?)(object?)textConnection)?.BackgroundQuery ?? false;
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                // PIA gap: WorkbookConnection.WebConnection is absent from the Excel PIA.
                dynamic? webConnection = null;
                try
                {
                    // PIA gap: WorkbookConnection has no WebConnection or root ConnectionString member.
                    webConnection = ((dynamic)(object)conn).WebConnection;
                    return webConnection?.BackgroundQuery ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException ex) when (IsUnsupportedConnectionProperty(ex))
        {
        }

        return false;
    }

    private static bool GetRefreshOnFileOpenSetting(Excel.WorkbookConnection conn)
    {
        try
        {
            int connType = (int)conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection?.RefreshOnFileOpen ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.RefreshOnFileOpen ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                Excel.TextConnection? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    // PIA gap: TextConnection.RefreshOnFileOpen is absent from the Excel PIA.
                    return ((dynamic?)(object?)textConnection)?.RefreshOnFileOpen ?? false;
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                // PIA gap: WorkbookConnection.WebConnection is absent from the Excel PIA.
                dynamic? webConnection = null;
                try
                {
                    // PIA gap: WorkbookConnection has no WebConnection or root ConnectionString member.
                    webConnection = ((dynamic)(object)conn).WebConnection;
                    return webConnection?.RefreshOnFileOpen ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException ex) when (IsUnsupportedConnectionProperty(ex))
        {
        }

        return false;
    }

    private static bool GetSavePasswordSetting(Excel.WorkbookConnection conn)
    {
        try
        {
            int connType = (int)conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection?.SavePassword ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.SavePassword ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                Excel.TextConnection? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    // PIA gap: TextConnection.SavePassword is absent from the Excel PIA.
                    return ((dynamic?)(object?)textConnection)?.SavePassword ?? false;
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                // PIA gap: WorkbookConnection.WebConnection is absent from the Excel PIA.
                dynamic? webConnection = null;
                try
                {
                    // PIA gap: WorkbookConnection has no WebConnection or root ConnectionString member.
                    webConnection = ((dynamic)(object)conn).WebConnection;
                    return webConnection?.SavePassword ?? false;
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException ex) when (IsUnsupportedConnectionProperty(ex))
        {
        }

        return false;
    }

    private static int GetRefreshPeriod(Excel.WorkbookConnection conn)
    {
        try
        {
            int connType = (int)conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection?.RefreshPeriod ?? 0;
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.RefreshPeriod ?? 0;
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                Excel.TextConnection? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    // PIA gap: TextConnection.RefreshPeriod is absent from the Excel PIA.
                    return ((dynamic?)(object?)textConnection)?.RefreshPeriod ?? 0;
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                // PIA gap: WorkbookConnection.WebConnection is absent from the Excel PIA.
                dynamic? webConnection = null;
                try
                {
                    // PIA gap: WorkbookConnection has no WebConnection or root ConnectionString member.
                    webConnection = ((dynamic)(object)conn).WebConnection;
                    return webConnection?.RefreshPeriod ?? 0;
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException ex) when (IsUnsupportedConnectionProperty(ex))
        {
        }

        return 0;
    }

    private static DateTime? GetLastRefreshDate(Excel.WorkbookConnection conn, ILogger logger)
    {
        try
        {
            int connType = (int)conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    // PIA gap: OLAP RefreshDate can be an empty VARIANT, not the PIA's non-nullable DateTime.
                    return oledbConnection == null ? null : GetRefreshDateSafe(((dynamic)oledbConnection).RefreshDate);
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return GetRefreshDateSafe(odbcConnection?.RefreshDate);
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }
        }
        catch (COMException ex)
        {
            logger.LogWarning(ex, "Could not read the connection's last-refresh timestamp; reporting it as unknown");
        }

        return null;
    }

    private static bool IsUnsupportedConnectionProperty(COMException exception) =>
        exception.HResult is unchecked((int)0x80020003) or unchecked((int)0x80004001);

    private static string? GetConnectionString(Excel.WorkbookConnection conn)
    {
        try
        {
            int connType = (int)conn.Type;
            string? connectionString = null;

            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    connectionString = oledbConnection?.Connection?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }
            else if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    connectionString = odbcConnection?.Connection?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }
            else if (connType is 3 or 4)
            {
                Excel.TextConnection? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    connectionString = textConnection?.Connection?.ToString();
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    // PIA gap: WorkbookConnection.WebConnection is absent from the Excel PIA.
                    dynamic? webConnection = null;
                    try
                    {
                        // PIA gap: WorkbookConnection has no WebConnection or root ConnectionString member.
                        webConnection = ((dynamic)(object)conn).WebConnection;
                        connectionString = webConnection?.Connection?.ToString();
                    }
                    catch (COMException)
                    {
                    }
                    finally
                    {
                        ComUtilities.Release(ref webConnection);
                    }
                }
            }

            // Fallback to root ConnectionString property
            if (string.IsNullOrWhiteSpace(connectionString))
            {
                try
                {
                    // PIA gap: WorkbookConnection has no WebConnection or root ConnectionString member.
                    connectionString = ((dynamic)(object)conn).ConnectionString?.ToString();
                }
                catch (COMException)
                {
                    // Property not available
                }
            }

            return connectionString;
        }
        catch (COMException)
        {
            // Property not available
        }

        return null;
    }

    private static string? GetCommandText(Excel.WorkbookConnection conn)
    {
        try
        {
            int connType = (int)conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    return oledbConnection?.CommandText?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }

            if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    return odbcConnection?.CommandText?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }

            if (connType is 3 or 4)
            {
                Excel.TextConnection? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection;
                    // PIA gap: TextConnection.CommandText is absent from the Excel PIA.
                    return ((dynamic?)(object?)textConnection)?.CommandText?.ToString();
                }
                catch (COMException)
                {
                }
                finally
                {
                    ComUtilities.Release(ref textConnection);
                }

                // PIA gap: WorkbookConnection.WebConnection is absent from the Excel PIA.
                dynamic? webConnection = null;
                try
                {
                    // PIA gap: WorkbookConnection has no WebConnection or root ConnectionString member.
                    webConnection = ((dynamic)(object)conn).WebConnection;
                    return webConnection?.CommandText?.ToString();
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
        }
        catch (COMException)
        {
        }

        return null;
    }

    private static string? GetCommandType(Excel.WorkbookConnection conn)
    {
        int commandType;
        try
        {
            int connType = (int)conn.Type;
            if (connType == 1)
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    if (oledbConnection == null)
                    {
                        return null;
                    }

                    commandType = Convert.ToInt32(oledbConnection.CommandType, CultureInfo.InvariantCulture);
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }
            else if (connType == 2)
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    if (odbcConnection == null)
                    {
                        return null;
                    }

                    commandType = Convert.ToInt32(odbcConnection.CommandType, CultureInfo.InvariantCulture);
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }
            else
            {
                return null;
            }
        }
        catch (COMException)
        {
            return null;
        }

        return commandType switch
        {
            1 => "Cube",
            2 => "SQL",
            3 => "Table",
            4 => "Default",
            5 => "List",
            _ => $"Unknown({commandType.ToString(CultureInfo.InvariantCulture)})"
        };
    }

    private static object GetConnectionProperties(Excel.WorkbookConnection conn, ILogger logger)
    {
        return new
        {
            BackgroundQuery = GetBackgroundQuerySetting(conn),
            RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn),
            SavePassword = GetSavePasswordSetting(conn),
            RefreshPeriod = GetRefreshPeriod(conn),
            LastRefresh = GetLastRefreshDate(conn, logger)
        };
    }

    private static void CreateConnection(Excel.Workbook workbook, string connectionName, ConnectionDefinition definition)
    {
        // Validate required fields
        if (string.IsNullOrWhiteSpace(definition.ConnectionString))
        {
            throw new InvalidOperationException("ConnectionString is required to create a connection.");
        }

        // Reject TEXT/WEB connection strings (legacy, use Power Query or ODC import instead)
        string connStr = definition.ConnectionString.Trim();
        if (connStr.StartsWith("TEXT;", StringComparison.OrdinalIgnoreCase) ||
            connStr.StartsWith("URL;", StringComparison.OrdinalIgnoreCase))
        {
            throw new NotSupportedException(
                "TEXT and WEB connections are no longer supported via create action. " +
                "Use powerquery tool for file/web imports, or create an ODC file and use import-odc action.");
        }

        Excel.Connections? connections = null;
        Excel.WorkbookConnection? newConnection = null;

        try
        {
            connections = workbook.Connections;

            object commandTypeArgument = DetermineCommandType(definition);

            // Use Add2() method (Add() is deprecated per Microsoft docs)
            // https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.connections.add2
            newConnection = connections.Add2(
                Name: connectionName,
                Description: definition.Description ?? "",
                ConnectionString: definition.ConnectionString,
                CommandText: definition.CommandText ?? "",
                lCmdtype: commandTypeArgument,
                CreateModelConnection: false,         // Don't create PowerPivot model connection
                ImportRelationships: false            // Don't import relationships
            );

            // Connection created successfully - let exceptions propagate naturally
        }
        finally
        {
            ComUtilities.Release(ref newConnection);
            ComUtilities.Release(ref connections);
        }
    }

    private static bool IsOlapConnection(Excel.WorkbookConnection connection)
    {
        int connectionType = Convert.ToInt32(connection.Type, CultureInfo.InvariantCulture);
        if (connectionType != 1) return false;

        Excel.OLEDBConnection? oledb = null;
        try
        {
            oledb = connection.OLEDBConnection;
            return oledb.OLAP;
        }
        finally
        {
            ComUtilities.Release(ref oledb);
        }
    }

    private static void UpdateConnectionProperties(Excel.WorkbookConnection conn, ConnectionDefinition definition)
    {
        int connType = Convert.ToInt32(conn.Type, CultureInfo.InvariantCulture);
        if (connType == 1 && definition.BackgroundQuery.HasValue &&
            !string.IsNullOrWhiteSpace(definition.ConnectionString))
        {
            string current;
            Excel.OLEDBConnection? providerConnection = null;
            try
            {
                providerConnection = conn.OLEDBConnection;
                current = Convert.ToString(providerConnection.Connection, CultureInfo.InvariantCulture)
                    ?? throw new InvalidOperationException("The OLEDB connection string could not be read.");
            }
            finally
            {
                ComUtilities.Release(ref providerConnection);
            }
            static DbConnectionStringBuilder Parse(string value)
            {
                value = value.Trim();
                return new DbConnectionStringBuilder
                {
                    ConnectionString = value.StartsWith("OLEDB;", StringComparison.OrdinalIgnoreCase) ? value[6..] : value
                };
            }
            var before = Parse(current);
            var after = Parse(definition.ConnectionString);
            before.TryGetValue("Provider", out var currentProvider);
            after.TryGetValue("Provider", out var targetProvider);
            if (!string.Equals(
                Convert.ToString(currentProvider, CultureInfo.InvariantCulture),
                Convert.ToString(targetProvider, CultureInfo.InvariantCulture), StringComparison.OrdinalIgnoreCase))
            {
                throw new ArgumentException(
                    "Change the OLEDB provider and backgroundQuery in separate calls. " +
                    "No connection properties have been changed.", nameof(definition));
            }
        }
        bool isOlap = IsOlapConnection(conn);
        if (isOlap && definition.BackgroundQuery == true)
        {
            throw new ArgumentException(
                "OLAP connections refresh synchronously and do not support backgroundQuery=true.",
                nameof(definition));
        }

        try
        {
            // Update description
            if (!string.IsNullOrWhiteSpace(definition.Description))
            {
                conn.Description = definition.Description;
            }

            if (connType == 1) // OLEDB
            {
                Excel.OLEDBConnection? oledbConnection = null;
                try
                {
                    oledbConnection = conn.OLEDBConnection;
                    if (oledbConnection != null)
                    {
                        if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                        {
                            oledbConnection.Connection = definition.ConnectionString;
                        }
                        if (!string.IsNullOrWhiteSpace(definition.CommandText))
                        {
                            oledbConnection.CommandText = definition.CommandText;
                        }
                        if (definition.BackgroundQuery.HasValue && !isOlap)
                        {
                            oledbConnection.BackgroundQuery = definition.BackgroundQuery.Value;
                        }

                        if (definition.RefreshOnFileOpen.HasValue)
                        {
                            oledbConnection.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                        }
                        if (definition.SavePassword.HasValue)
                        {
                            oledbConnection.SavePassword = definition.SavePassword.Value;
                        }
                        if (definition.RefreshPeriod.HasValue)
                        {
                            oledbConnection.RefreshPeriod = definition.RefreshPeriod.Value;
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref oledbConnection);
                }
            }
            else if (connType == 2) // ODBC
            {
                Excel.ODBCConnection? odbcConnection = null;
                try
                {
                    odbcConnection = conn.ODBCConnection;
                    if (odbcConnection != null)
                    {
                        if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                        {
                            odbcConnection.Connection = definition.ConnectionString;
                        }
                        if (!string.IsNullOrWhiteSpace(definition.CommandText))
                        {
                            odbcConnection.CommandText = definition.CommandText;
                        }
                        if (definition.BackgroundQuery.HasValue)
                        {
                            odbcConnection.BackgroundQuery = definition.BackgroundQuery.Value;
                        }
                        if (definition.RefreshOnFileOpen.HasValue)
                        {
                            odbcConnection.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                        }
                        if (definition.SavePassword.HasValue)
                        {
                            odbcConnection.SavePassword = definition.SavePassword.Value;
                        }
                        if (definition.RefreshPeriod.HasValue)
                        {
                            odbcConnection.RefreshPeriod = definition.RefreshPeriod.Value;
                        }
                    }
                }
                finally
                {
                    ComUtilities.Release(ref odbcConnection);
                }
            }
            else if (connType is 3 or 4) // TEXT (type 3) or WEB (type 4) - Excel may report CSV files as either
            {
                // Excel has type 3/4 confusion: CSV files created with "TEXT;filepath" may be reported as type 4 (WEB)
                // Try TextConnection first (correct for type 3), fall back to WebConnection if that fails
                Excel.TextConnection? textConnection = null;
                try
                {
                    textConnection = conn.TextConnection; // Try TEXT first
                }
                catch (COMException)
                {
                }

                if (textConnection != null)
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                        {
                            textConnection.Connection = definition.ConnectionString;
                        }
                        if (!string.IsNullOrWhiteSpace(definition.CommandText))
                        {
                            // PIA gap: TextConnection.CommandText is absent from the Excel PIA.
                            ((dynamic)(object)textConnection).CommandText = definition.CommandText;
                        }
                        if (definition.BackgroundQuery.HasValue)
                        {
                            // PIA gap: TextConnection.BackgroundQuery is absent from the Excel PIA.
                            ((dynamic)(object)textConnection).BackgroundQuery = definition.BackgroundQuery.Value;
                        }
                        if (definition.RefreshOnFileOpen.HasValue)
                        {
                            // PIA gap: TextConnection.RefreshOnFileOpen is absent from the Excel PIA.
                            ((dynamic)(object)textConnection).RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                        }
                        if (definition.SavePassword.HasValue)
                        {
                            // PIA gap: TextConnection.SavePassword is absent from the Excel PIA.
                            ((dynamic)(object)textConnection).SavePassword = definition.SavePassword.Value;
                        }
                        if (definition.RefreshPeriod.HasValue)
                        {
                            // PIA gap: TextConnection.RefreshPeriod is absent from the Excel PIA.
                            ((dynamic)(object)textConnection).RefreshPeriod = definition.RefreshPeriod.Value;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref textConnection);
                    }

                    return;
                }

                // PIA gap: WorkbookConnection.WebConnection is absent from the Excel PIA.
                dynamic? webConnection = null;
                try
                {
                    // PIA gap: WorkbookConnection has no WebConnection or root ConnectionString member.
                    webConnection = ((dynamic)(object)conn).WebConnection; // Fall back to WEB
                    if (webConnection != null)
                    {
                        if (!string.IsNullOrWhiteSpace(definition.ConnectionString))
                        {
                            webConnection.Connection = definition.ConnectionString;
                        }
                        if (!string.IsNullOrWhiteSpace(definition.CommandText))
                        {
                            webConnection.CommandText = definition.CommandText;
                        }
                        if (definition.BackgroundQuery.HasValue)
                        {
                            webConnection.BackgroundQuery = definition.BackgroundQuery.Value;
                        }
                        if (definition.RefreshOnFileOpen.HasValue)
                        {
                            webConnection.RefreshOnFileOpen = definition.RefreshOnFileOpen.Value;
                        }
                        if (definition.SavePassword.HasValue)
                        {
                            webConnection.SavePassword = definition.SavePassword.Value;
                        }
                        if (definition.RefreshPeriod.HasValue)
                        {
                            webConnection.RefreshPeriod = definition.RefreshPeriod.Value;
                        }
                    }
                }
                catch (COMException)
                {
                    // Neither works - skip property updates
                }
                finally
                {
                    ComUtilities.Release(ref webConnection);
                }
            }
            else if (connType == 5) // XMLMAP
            {
                // XMLMAP connection properties - future implementation
                // For now, just update basic properties like description (already done above)
            }
            else if (connType == 6) // DATAFEED
            {
                // DATAFEED connection properties - future implementation
            }
            else if (connType == 7) // MODEL
            {
                // MODEL connection properties - future implementation
            }
            else if (connType == 8) // WORKSHEET
            {
                // WORKSHEET connection properties - future implementation
            }
            else if (connType == 9) // NOSOURCE
            {
                // NOSOURCE connection properties - future implementation
            }
            else
            {
                // Unknown connection type - skip property updates
                // Description was already updated above if provided
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Failed to update connection properties: {ex.Message}", ex);
        }
    }

    private static object DetermineCommandType(ConnectionDefinition definition)
    {
        if (!string.IsNullOrWhiteSpace(definition.CommandType))
        {
            var value = definition.CommandType.Trim();
            if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var numeric))
            {
                return numeric;
            }

            return value.ToLowerInvariant() switch
            {
                "cube" => 1,
                "sql" => 2,
                "table" => 3,
                "default" => 4,
                "list" => 5,
                _ => Type.Missing
            };
        }

        if (!string.IsNullOrWhiteSpace(definition.CommandText))
        {
            // When command text is provided we default to SQL command type (2).
            return 2;
        }

        return Type.Missing;
    }


    private static void CreateQueryTableForConnection(
        Excel.Worksheet targetSheet,
        Excel.WorkbookConnection conn,
        PowerQueryHelpers.QueryTableOptions options,
        CancellationToken cancellationToken)
    {
        // For regular connections (not Power Query), we need connection string
        string? connectionString = GetConnectionString(conn);
        string? commandText = GetCommandText(conn);

        if (string.IsNullOrWhiteSpace(connectionString))
        {
            throw new InvalidOperationException("Connection has no connection string");
        }

        // Command text can be empty for some connection types (Text, Web)
        // Use empty string if not provided
        if (string.IsNullOrWhiteSpace(commandText))
        {
            commandText = "";
        }

        Excel.QueryTables? queryTables = null;
        Excel.QueryTable? queryTable = null;
        Excel.Range? range = null;

        try
        {
            queryTables = targetSheet.QueryTables;
            range = targetSheet.Range["A1"];
            queryTable = queryTables.Add(conn, range, commandText);

            queryTable.Name = options.Name.Replace(" ", "_");
            queryTable.RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells;
            queryTable.BackgroundQuery = options.BackgroundQuery;
            queryTable.RefreshOnFileOpen = options.RefreshOnFileOpen;
            queryTable.SavePassword = options.SavePassword;
            queryTable.PreserveColumnInfo = options.PreserveColumnInfo;
            queryTable.PreserveFormatting = options.PreserveFormatting;
            queryTable.AdjustColumnWidth = options.AdjustColumnWidth;

            if (options.RefreshImmediately)
            {
                // Do NOT use EnterLongOperation here. Synchronous QueryTable refresh can depend on
                // inbound Excel callbacks to complete, and rejecting them can deadlock the load.
                QueryTableRefreshHelper.RefreshSynchronously(
                    queryTable,
                    cancellationToken,
                    $"Connection load for '{options.Name}'");
            }
        }
        finally
        {
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref queryTable);
            ComUtilities.Release(ref queryTables);
        }
    }

    #endregion
}

/// <summary>
/// Connection definition for JSON import/export
/// </summary>
internal sealed class ConnectionDefinition
{
    public string Name { get; set; } = "";
    public string? Description { get; set; }
    public string Type { get; set; } = "";
    public string? ConnectionString { get; set; }
    public string? CommandText { get; set; }
    public string? CommandType { get; set; }
    public bool? BackgroundQuery { get; set; }
    public bool? RefreshOnFileOpen { get; set; }
    public bool? SavePassword { get; set; }
    public int? RefreshPeriod { get; set; }
}
