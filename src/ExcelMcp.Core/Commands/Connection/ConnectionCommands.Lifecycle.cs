using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.CSharp.RuntimeBinder;
using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Utilities;
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
            Excel.Connections? connections = null;

            try
            {
                connections = ctx.Book.Connections;

                for (int i = 1; i <= connections.Count; i++)
                {
                    Excel.WorkbookConnection? conn = null;
                    try
                    {
                        conn = connections.Item(i);

                        var connInfo = new ConnectionInfo
                        {
                            Name = conn.Name?.ToString() ?? "",
                            Description = conn.Description?.ToString() ?? "",
                            Type = ConnectionHelpers.GetConnectionTypeName((int)conn.Type),
                            IsPowerQuery = PowerQueryHelpers.IsPowerQueryConnection(conn),
                            BackgroundQuery = GetBackgroundQuerySetting(conn),
                            RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn),
                            LastRefresh = GetLastRefreshDate(conn, batch.Logger)
                        };

                        result.Connections.Add(connInfo);
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
            Excel.WorkbookConnection? conn = null;
            try
            {
                conn = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName);

                if (conn == null)
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                }

                result.Type = ConnectionHelpers.GetConnectionTypeName((int)conn.Type);
                result.IsPowerQuery = PowerQueryHelpers.IsPowerQueryConnection(conn);

                string sanitizedConnectionString =
                    ConnectionStringSanitizer.Sanitize(GetConnectionString(conn)) ?? "";
                result.ConnectionString = sanitizedConnectionString;

                result.CommandText = GetCommandText(conn);
                result.CommandType = GetCommandType(conn);

                var definition = new
                {
                    Name = connectionName,
                    Type = result.Type,
                    Description = conn.Description?.ToString() ?? "",
                    IsPowerQuery = result.IsPowerQuery,
                    ConnectionString = sanitizedConnectionString,
                    CommandText = result.CommandText,
                    CommandType = result.CommandType,
                    Properties = GetConnectionProperties(conn, batch.Logger)
                };

                result.DefinitionJson = JsonSerializer.Serialize(definition, s_jsonOptions);

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref conn);
            }
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
            Excel.WorkbookConnection? conn = null;
            try
            {
                conn = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName);

                if (conn == null)
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                }

                if (PowerQueryHelpers.IsPowerQueryConnection(conn))
                {
                    if (PowerQueryHelpers.IsOrphanedPowerQueryConnection(ctx.Book, conn))
                    {
                        throw new InvalidOperationException($"Connection '{connectionName}' is an orphaned Power Query connection with no corresponding query. Use connection 'delete' to remove it.");
                    }
                    throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Use powerquery 'refresh' instead.");
                }

                RefreshWorkbookConnection(conn, ct);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref conn);
            }
        }, timeoutCts.Token);  // Extended timeout (default 5 minutes) for slow data sources
    }

    [System.Diagnostics.CodeAnalysis.SuppressMessage("ErrorHandling", "EXL002",
        Justification = "Retains refresh and BackgroundQuery restoration errors through COM cleanup, then rethrows the original error or both failures; cancellation is preserved.")]
    internal static void RefreshWorkbookConnection(Excel.WorkbookConnection connection, CancellationToken cancellationToken)
    {
        Excel.OLEDBConnection? oledbConnection = null;
        Excel.ODBCConnection? odbcConnection = null;
        bool originalBackgroundQuery = false;
        bool canRestoreBackgroundQuery = false;
        Exception? refreshFailure = null;
        COMException? restoreFailure = null;

        try
        {
            try
            {
                // Synchronous refresh avoids waiting on Excel's background callbacks.
                if (connection.Type == Excel.XlConnectionType.xlConnectionTypeOLEDB)
                {
                    oledbConnection = connection.OLEDBConnection;
                    if (oledbConnection is not null && !oledbConnection.OLAP)
                    {
                        originalBackgroundQuery = oledbConnection.BackgroundQuery;
                        canRestoreBackgroundQuery = true;
                        oledbConnection.BackgroundQuery = false;
                    }
                }
                else if (connection.Type == Excel.XlConnectionType.xlConnectionTypeODBC)
                {
                    odbcConnection = connection.ODBCConnection;
                    if (odbcConnection is not null)
                    {
                        originalBackgroundQuery = odbcConnection.BackgroundQuery;
                        canRestoreBackgroundQuery = true;
                        odbcConnection.BackgroundQuery = false;
                    }
                }
            }
            catch (COMException ex) when (ComErrors.IsMemberUnavailable(ex))
            {
                // Provider doesn't support BackgroundQuery — proceed with default behavior.
            }

            // Do NOT use EnterLongOperation here. Like Power Query refresh, synchronous
            // connection.Refresh() can require inbound Excel/provider callbacks to complete.
            // Rejecting those callbacks can deadlock the refresh.
            OleMessageFilter.SetPendingCancellationToken(cancellationToken);
            try
            {
                if (connection.Type is Excel.XlConnectionType.xlConnectionTypeTEXT or Excel.XlConnectionType.xlConnectionTypeWEB)
                    RefreshLegacyQueryTables(connection, cancellationToken);
                else
                    connection.Refresh();
            }
            finally
            {
                OleMessageFilter.ClearPendingCancellationToken();
            }

            if (ReadRefreshStatus(connection).Supported)
            {
                WaitForConnectionRefreshCompletion(
                    () => ReadRefreshStatus(connection).Refreshing,
                    () => CancelTypedRefresh(connection, ""),
                    cancellationToken);
            }
        }
        catch (Exception ex)
        {
            refreshFailure = ex;
        }
        finally
        {
            try
            {
                if (canRestoreBackgroundQuery)
                {
                    if (oledbConnection is not null)
                        oledbConnection.BackgroundQuery = originalBackgroundQuery;
                    else if (odbcConnection is not null)
                        odbcConnection.BackgroundQuery = originalBackgroundQuery;
                }
            }
            catch (COMException ex)
            {
                restoreFailure = ex;
            }
            finally
            {
                ComUtilities.Release(ref odbcConnection);
                ComUtilities.Release(ref oledbConnection);
            }
        }
        if (restoreFailure != null)
        {
            var failure = new InvalidOperationException(
                "Could not restore the connection's original BackgroundQuery setting after the refresh attempt. " +
                "Inspect the connection before retrying.",
                refreshFailure == null ? restoreFailure : new AggregateException(refreshFailure, restoreFailure));
            if (refreshFailure is OperationCanceledException canceled)
                throw new OperationCanceledException(failure.Message, failure, canceled.CancellationToken);
            throw failure;
        }
        if (refreshFailure != null)
        {
            System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(refreshFailure).Throw();
        }
    }

    private static void RefreshLegacyQueryTables(Excel.WorkbookConnection connection, CancellationToken cancellationToken)
    {
        Excel.Ranges? ranges = null;
        try
        {
            ranges = connection.Ranges;
            for (int index = 1; index <= ranges.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                Excel.Range? range = null;
                Excel.QueryTable? query = null;
                try
                {
                    range = ranges.Item[index];
                    query = range.QueryTable;
                    if (!query.Refresh(false))
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        throw new InvalidOperationException($"Excel did not complete refresh of legacy connection '{connection.Name}'.");
                    }
                    cancellationToken.ThrowIfCancellationRequested();
                }
                finally
                {
                    ComUtilities.Release(ref query);
                    ComUtilities.Release(ref range);
                }
            }
            if (ranges.Count == 0)
                connection.Refresh();
        }
        finally
        {
            ComUtilities.Release(ref ranges);
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
            cancellationToken.ThrowIfCancellationRequested();
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
        catch (OperationCanceledException ex)
        {
            try
            {
                cancelRefresh();
            }
            catch (Exception cancelFailure) when (cancelFailure is COMException or RuntimeBinderException)
            {
                throw new OperationCanceledException(
                    "Refresh was cancelled or timed out, but Excel could not confirm cancellation. " +
                    "The query may still be running; inspect refresh status before saving or closing.",
                    new AggregateException(ex, cancelFailure), cancellationToken);
            }
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
            Excel.WorkbookConnection? conn = null;
            try
            {
                conn = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName);

                if (conn == null)
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                }

                if (PowerQueryHelpers.IsPowerQueryConnection(conn) &&
                    !PowerQueryHelpers.IsOrphanedPowerQueryConnection(ctx.Book, conn))
                {
                    throw new InvalidOperationException(
                        $"Connection '{connectionName}' is a Power Query connection. Use powerquery with action 'delete' instead.");
                }

                PowerQueryHelpers.RemoveQueryTables(ctx.Book, conn);
                ComUtilities.Release(ref conn);
                conn = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName);
                conn?.Delete();
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref conn);
            }
        });
    }
}
