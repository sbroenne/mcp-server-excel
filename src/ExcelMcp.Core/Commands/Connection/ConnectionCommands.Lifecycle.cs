using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection lifecycle operations (List, View, Import, Export, Update, Delete)
/// </summary>
public partial class ConnectionCommands
{
    private static readonly JsonSerializerOptions s_jsonOptions = new() { WriteIndented = true };

    /// <summary>
    /// Lists all connections in a workbook
    /// </summary>
    public ConnectionListResult List(IExcelBatch batch)
    {
        var result = new ConnectionListResult { FilePath = batch.WorkbookPath };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? connections = null;

            try
            {
                connections = ctx.Book.Connections;

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
                    catch (System.Runtime.InteropServices.COMException)
                    {
                        // Skip connections that have COM access issues
                        continue;
                    }
                    finally
                    {
                        ComUtilities.Release(ref conn);
                    }
                }

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref connections);
            }
        });
    }

    /// <summary>
    /// Views detailed connection information
    /// </summary>
    public ConnectionViewResult View(IExcelBatch batch, string connectionName)
    {
        var result = new ConnectionViewResult
        {
            FilePath = batch.WorkbookPath,
            ConnectionName = connectionName
        };

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                throw new InvalidOperationException($"Connection '{connectionName}' not found.");
            }

            result.Type = ConnectionHelpers.GetConnectionTypeName((int)conn.Type);
            result.IsPowerQuery = PowerQueryHelpers.IsPowerQueryConnection(conn);

            // Get connection string (raw for LLM usage - sanitization removed)
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
            return result;
        });
    }

    /// <summary>
    /// Creates a new connection in the workbook
    /// </summary>
    public OperationResult Create(IExcelBatch batch, string connectionName,
        string connectionString, string? commandText = null, string? description = null)
    {
        return batch.Execute((ctx, ct) =>
        {
            // Create connection definition
            var definition = new ConnectionDefinition
            {
                Name = connectionName,
                Description = description ?? "",
                ConnectionString = connectionString,
                CommandText = commandText ?? "",
                CommandType = string.IsNullOrWhiteSpace(commandText) ? null : "SQL",
                SavePassword = false // Default to secure setting
            };

            // Create the connection using existing helper method
            CreateConnection(ctx.Book, connectionName, definition);
            return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
        });
    }

    /// <summary>
    /// Refreshes connection data
    /// </summary>
    public OperationResult Refresh(IExcelBatch batch, string connectionName)
    {
        return Refresh(batch, connectionName, timeout: null);
    }

    /// <summary>
    /// Refreshes connection data with timeout
    /// </summary>
    public OperationResult Refresh(IExcelBatch batch, string connectionName, TimeSpan? timeout)
    {
        var effectiveTimeout = timeout ?? ComInteropConstants.DataOperationTimeout;
        using var timeoutCts = new CancellationTokenSource(effectiveTimeout);

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                throw new InvalidOperationException($"Connection '{connectionName}' not found.");
            }

            // Check if this is a Power Query connection (handle separately)
            if (PowerQueryHelpers.IsPowerQueryConnection(conn))
            {
                // Check if this is an orphaned Power Query connection
                if (PowerQueryHelpers.IsOrphanedPowerQueryConnection(ctx.Book, conn))
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' is an orphaned Power Query connection with no corresponding query. Use connection 'delete' to remove it.");
                }
                throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Use powerquery 'refresh' instead.");
            }

            RefreshWorkbookConnection(conn, ct);
            return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
        }, timeoutCts.Token);  // Extended timeout (default 5 minutes) for slow data sources
    }

    private static void RefreshWorkbookConnection(Excel.WorkbookConnection connection, CancellationToken cancellationToken)
    {
        dynamic? subConnection = null;
        bool originalBackgroundQuery = false;
        bool canRestoreBackgroundQuery = false;
        bool supportsRefreshing = false;

        try
        {
            try
            {
                subConnection = GetTypedSubConnection(connection);
                if (subConnection != null)
                {
                    originalBackgroundQuery = subConnection.BackgroundQuery;
                    canRestoreBackgroundQuery = true;

                    // CRITICAL: Force BackgroundQuery = false to ensure synchronous refresh.
                    //
                    // With BackgroundQuery = true (async), connection.Refresh() returns immediately
                    // while Excel processes the query in a background thread. We then poll
                    // connection.Refreshing with Thread.Sleep(5000). On STA threads with the
                    // OleMessageFilter registered, COM events from Excel during the background refresh
                    // cause Thread.Sleep to return via MsgWaitForMultipleObjectsEx — turning the
                    // polling loop into a 100% CPU spin for the entire duration of the refresh.
                    //
                    // With BackgroundQuery = false (synchronous), connection.Refresh() blocks the
                    // STA thread until done. connection.Refreshing is false when it returns, so
                    // WaitForConnectionRefreshCompletion exits immediately with zero CPU overhead.
                    subConnection.BackgroundQuery = false;
                }
            }
            catch (COMException)
            {
                // Provider doesn't support BackgroundQuery — proceed with default behavior.
            }
            catch (RuntimeBinderException)
            {
                // Sub-connection doesn't expose BackgroundQuery — proceed with default behavior.
            }

            // Enter long operation mode: MessagePending returns WAITDEFPROCESS to dispatch
            // to HandleInComingCall, which rejects with SERVERCALL_RETRYLATER.
            // This triggers the caller's RetryRejectedCall backoff instead of either:
            // - WAITNOPROCESS rejection storm (88% CPU) or
            // - WAITDEFPROCESS + EnsureScanDefinedEvents spin (97% CPU)
            OleMessageFilter.EnterLongOperation();
            try
            {
                connection.Refresh();
            }
            finally
            {
                OleMessageFilter.ExitLongOperation();
            }

            try
            {
                // PIA gap: WorkbookConnection.Refreshing not in Microsoft.Office.Interop.Excel v16 PIA
                _ = ((dynamic)connection).Refreshing;
                supportsRefreshing = true;
            }
            catch (COMException)
            {
                supportsRefreshing = false;
            }
            catch (RuntimeBinderException)
            {
                supportsRefreshing = false;
            }

            if (supportsRefreshing)
            {
                WaitForConnectionRefreshCompletion(
                    () =>
                    {
                        try
                        {
                            // PIA gap: WorkbookConnection.Refreshing not in Microsoft.Office.Interop.Excel v16 PIA
                            return ((dynamic)connection).Refreshing;
                        }
                        catch (COMException)
                        {
                            return false;
                        }
                        catch (RuntimeBinderException)
                        {
                            return false;
                        }
                    },
                    () =>
                    {
                        try
                        {
                            // PIA gap: WorkbookConnection.CancelRefresh not in Microsoft.Office.Interop.Excel v16 PIA
                            ((dynamic)connection).CancelRefresh();
                        }
                        catch (COMException)
                        {
                            // Provider does not support cancellation.
                        }
                        catch (RuntimeBinderException)
                        {
                            // Provider does not expose cancellation.
                        }
                    },
                    cancellationToken);
            }
        }
        finally
        {
            if (canRestoreBackgroundQuery && subConnection != null)
            {
                try
                {
                    subConnection.BackgroundQuery = originalBackgroundQuery;
                }
                catch (COMException)
                {
                    // Ignore inability to restore provider-specific setting.
                }
            }

            ComUtilities.Release(ref subConnection);
        }
    }

    private static void WaitForConnectionRefreshCompletion(
        Func<bool> isRefreshing,
        Action cancelRefresh,
        CancellationToken cancellationToken)
    {
        // CRITICAL: Rate-limit the isRefreshing() COM call to every 5000ms of *real* elapsed time.
        // See WaitForRefreshCompletion in PowerQueryCommands.Helpers.cs for full explanation.
        const int CheckIntervalMs = 5000;
        var sw = System.Diagnostics.Stopwatch.StartNew();
        try
        {
            if (!isRefreshing())
                return;

            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                // Sleep without pumping the STA COM queue. See WaitForRefreshCompletion for details.
                ComUtilities.KernelSleep(CheckIntervalMs);
                if (sw.Elapsed.TotalMilliseconds < CheckIntervalMs)
                    continue;
                sw.Restart();
                if (!isRefreshing())
                    break;
            }
        }
        catch (OperationCanceledException)
        {
            cancelRefresh();
            throw;
        }
    }

    /// <summary>
    /// Deletes a connection
    /// </summary>
    public OperationResult Delete(IExcelBatch batch, string connectionName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                throw new InvalidOperationException($"Connection '{connectionName}' not found.");
            }

            // Check if this is a Power Query connection
            if (PowerQueryHelpers.IsPowerQueryConnection(conn))
            {
                // Check if this is an orphaned Power Query connection (no corresponding query exists)
                // Orphaned connections can be safely deleted via the connection API
                if (!PowerQueryHelpers.IsOrphanedPowerQueryConnection(ctx.Book, conn))
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Use powerquery with action 'Delete' instead.");
                }
                // Orphaned connection - allow deletion to proceed
            }

            // Remove associated QueryTables first
            PowerQueryHelpers.RemoveQueryTables(ctx.Book, connectionName);

            // Delete the connection
            conn.Delete();
            return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
        });
    }
}



