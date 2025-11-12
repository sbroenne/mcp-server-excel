using System.Text.Json;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection commands - FilePath-based API implementations
/// Simple read operations converted to use FileHandleManager pattern
/// </summary>
public partial class ConnectionCommands
{
    /// <summary>
    /// Lists all connections in a workbook using FilePath API
    /// </summary>
    public async Task<ConnectionListResult> ListAsync(string filePath)
    {
        var result = new ConnectionListResult { FilePath = filePath };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? connections = null;

                try
                {
                    connections = handle.Workbook.Connections;

                    for (int i = 1; i <= connections.Count; i++)
                    {
                        dynamic? conn = null;
                        try
                        {
                            conn = connections.Item(i);

                            var connInfo = new ConnectionInfo
                            {
                                Name = conn.Name?.ToString() ?? "",
                                Description = conn.Description?.ToString() ?? "",
                                Type = ConnectionHelpers.GetConnectionTypeName(conn.Type),
                                IsPowerQuery = PowerQueryHelpers.IsPowerQueryConnection(conn),
                                BackgroundQuery = GetBackgroundQuerySetting(conn),
                                RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn),
                                LastRefresh = GetLastRefreshDate(conn)
                            };

                            result.Connections.Add(connInfo);
                        }
                        catch
                        {
                            // Skip connections that can't be read
                            continue;
                        }
                        finally
                        {
                            ComUtilities.Release(ref conn);
                        }
                    }

                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref connections);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error listing connections: {ex.Message}";
        }

        return result;
    }

    /// <summary>
    /// Views detailed connection information (FilePath-based API)
    /// </summary>
    public async Task<ConnectionViewResult> ViewAsync(string filePath, string connectionName)
    {
        var result = new ConnectionViewResult
        {
            FilePath = filePath,
            ConnectionName = connectionName
        };

        try
        {
            var handle = await FileHandleManager.Instance.OpenOrGetAsync(filePath);

            await Task.Run(() =>
            {
                dynamic? conn = ComUtilities.FindConnection(handle.Workbook, connectionName);

                if (conn == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' not found";
                    return;
                }

                try
                {
                    result.Type = ConnectionHelpers.GetConnectionTypeName(conn.Type);
                    result.IsPowerQuery = PowerQueryHelpers.IsPowerQueryConnection(conn);

                    // Get connection string
                    string? rawConnectionString = GetConnectionString(conn);
                    result.ConnectionString = rawConnectionString ?? "";

                    // Get command text and type
                    result.CommandText = GetCommandText(conn);
                    result.CommandType = GetCommandType(conn);

                    // Build comprehensive JSON definition
                    var definition = new
                    {
                        Name = connectionName,
                        Type = result.Type,
                        Description = conn.Description?.ToString() ?? "",
                        IsPowerQuery = result.IsPowerQuery,
                        ConnectionString = result.ConnectionString,
                        CommandText = result.CommandText,
                        CommandType = result.CommandType,
                        Properties = GetConnectionProperties(conn)
                    };

                    result.DefinitionJson = JsonSerializer.Serialize(definition, s_jsonOptions);
                    result.Success = true;
                }
                finally
                {
                    ComUtilities.Release(ref conn);
                }
            });
        }
        catch (Exception ex)
        {
            result.Success = false;
            result.ErrorMessage = $"Error viewing connection: {ex.Message}";
        }

        return result;
    }
}
