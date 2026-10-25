using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Sbroenne.ExcelMcp.Core.Utilities;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection operations (LoadTo, Test)
/// </summary>
public partial class ConnectionCommands
{
    /// <summary>
    /// Loads connection data to a worksheet
    /// </summary>
    public OperationResult LoadTo(IExcelBatch batch, string connectionName, string sheetName)
    {
        using var timeoutCts = new CancellationTokenSource(ComInteropConstants.DataOperationTimeout);

        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? conn = null;
            Excel.Sheets? sheets = null;
            Excel.Worksheet? targetSheet = null;
            Excel.QueryTable? legacyQuery = null;

            try
            {
                conn = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName);

                if (conn == null)
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                }

                // Check if this is a Power Query connection
                if (PowerQueryHelpers.IsPowerQueryConnection(conn))
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Use 'pq-loadto' command instead.");
                }

                if (conn.Type is Excel.XlConnectionType.xlConnectionTypeTEXT or Excel.XlConnectionType.xlConnectionTypeWEB)
                    legacyQuery = GetLegacyImportForLoad(conn);

                var definition = new ConnectionDefinition
                {
                    Name = connectionName,
                    Description = conn.Description,
                    ConnectionString = GetConnectionString(conn),
                    CommandText = GetCommandText(conn),
                    CommandType = GetCommandType(conn),
                    BackgroundQuery = GetBackgroundQuerySetting(conn),
                    RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn),
                    SavePassword = GetSavePasswordSetting(conn),
                    RefreshPeriod = GetRefreshPeriod(conn)
                };

                // Find or create target sheet
                sheets = ctx.Book.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    Excel.Worksheet? sheet = null;
                    try
                    {
                        sheet = (Excel.Worksheet)sheets[i];
                        if (sheet.Name.ToString().Equals(sheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            targetSheet = sheet;
                            sheet = null; // Don't release in finally since we're keeping reference
                            break;
                        }
                    }
                    finally
                    {
                        ComUtilities.Release(ref sheet);
                    }
                }

                if (targetSheet == null)
                {
                    targetSheet = (Excel.Worksheet)sheets.Add();
                    targetSheet.Name = sheetName;
                }

                if (legacyQuery != null)
                {
                    LoadLegacyImport(targetSheet, legacyQuery, connectionName, definition.Description, ct);
                    return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
                }

                // Remove existing QueryTables first
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, conn);
                ComUtilities.Release(ref conn);
                conn = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName);
                if (conn == null)
                {
                    CreateConnection(ctx.Book, connectionName, definition);
                    conn = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName)
                        ?? throw new InvalidOperationException(
                            $"Connection '{connectionName}' could not be recreated after replacing its QueryTables.");
                    UpdateConnectionProperties(conn, definition);
                }

                // Create QueryTable to load data
                var options = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = connectionName,
                    RefreshImmediately = true
                };

                CreateQueryTableForConnection(targetSheet, conn, options, ct);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref legacyQuery);
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref conn);
            }
        }, timeoutCts.Token);
    }

    private static Excel.QueryTable GetLegacyImportForLoad(Excel.WorkbookConnection connection)
    {
        Excel.Ranges? ranges = null;
        Excel.Range? range = null;
        try
        {
            ranges = connection.Ranges;
            if (ranges.Count != 1)
                throw new InvalidOperationException(
                    $"Connection '{connection.Name}' must have exactly one linked import to move it with its settings. " +
                    "No QueryTables have been changed.");
            range = ranges[1];
            return range.QueryTable
                ?? throw new InvalidOperationException($"Connection '{connection.Name}' has no linked import.");
        }
        finally
        {
            ComUtilities.Release(ref range);
            ComUtilities.Release(ref ranges);
        }
    }

    private static void LoadLegacyImport(
        Excel.Worksheet targetSheet, Excel.QueryTable source, string connectionName, string? description,
        CancellationToken cancellationToken)
    {
        if (targetSheet.ProtectContents && !targetSheet.ProtectionMode)
            throw new InvalidOperationException(
                $"Worksheet '{targetSheet.Name}' is protected. No QueryTables have been changed.");

        Excel.QueryTables? queries = null;
        Excel.Range? destination = null;
        Excel.QueryTable? loaded = null;
        Excel.WorkbookConnection? loadedConnection = null;
        Excel.Range? sourceDestination = null;
        Excel.Worksheet? sourceSheet = null;
        string? sourceName = null;
        string? sourceSheetName = null;
        string? copiedName = null;
        string? copiedConnectionName = null;
        string targetSheetName = targetSheet.Name;
        string sourceState = "The original import has not been deleted.";
        try
        {
            sourceDestination = source.Destination;
            sourceSheet = sourceDestination.Worksheet;
            sourceName = source.Name;
            sourceSheetName = sourceSheet.Name;
            if (sourceSheet.Name.Equals(targetSheet.Name, StringComparison.OrdinalIgnoreCase) &&
                sourceDestination.Row == 1 && sourceDestination.Column == 1)
            {
                QueryTableRefreshHelper.RefreshSynchronously(
                    source, cancellationToken, $"Connection load for '{connectionName}'");
                return;
            }

            queries = targetSheet.QueryTables;
            destination = targetSheet.Range["A1"];
            // Excel copies the existing import definition, including settings absent from WorkbookConnection.
            loaded = queries.Add(source, destination);
            loadedConnection = loaded.WorkbookConnection;
            copiedConnectionName = loadedConnection.Name;
            loaded.Name = connectionName.Replace(" ", "_");
            copiedName = loaded.Name;
            QueryTableRefreshHelper.RefreshSynchronously(
                loaded, cancellationToken, $"Connection load for '{connectionName}'");
            sourceState = "Deletion of the original import was requested but has not been confirmed.";
            source.Delete();
            sourceState = "The original import was deleted; this command did not clear its original cells.";
            loadedConnection.Name = connectionName;
            copiedConnectionName = connectionName;
            // The linked-import property path also preserves the copied description after reopening.
            UpdateConnectionProperties(loadedConnection, new ConnectionDefinition { Description = description });
        }
        catch (Exception exception) when (loaded is not null &&
            (CreatedObjectFailure.CanReport(exception) || exception is OperationCanceledException))
        {
            string copy = $"QueryTable '{copiedName ?? "(name not confirmed)"}' on sheet '{targetSheetName}' at A1";
            string connection = copiedConnectionName is null
                ? "whose connection name has not been confirmed"
                : $"with connection '{copiedConnectionName}'";
            var failure = CreatedObjectFailure.CreateDescribed(
                $"Moving connection '{connectionName}'",
                $"{copy} {connection}",
                $"Original QueryTable '{sourceName}' on sheet '{sourceSheetName}': {sourceState} " +
                $"The copied {copy} {connection} remains; its refresh or settings may be incomplete. " +
                "No cleanup or rollback was performed.",
                exception);
            throw exception switch
            {
                OperationCanceledException canceled =>
                    new OperationCanceledException(failure.Message, canceled, canceled.CancellationToken),
                TimeoutException timeout => new TimeoutException(failure.Message, timeout),
                _ => failure
            };
        }
        finally
        {
            ComUtilities.Release(ref loadedConnection);
            ComUtilities.Release(ref loaded);
            ComUtilities.Release(ref destination);
            ComUtilities.Release(ref queries);
            ComUtilities.Release(ref sourceSheet);
            ComUtilities.Release(ref sourceDestination);
        }
    }

    /// <summary>
    /// Gets connection properties
    /// </summary>

    public OperationResult Test(IExcelBatch batch, string connectionName)
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

                int connType = (int)conn.Type;

                if (connType is 4 or 5)
                {
                    return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
                }

                string? connectionString = GetConnectionString(conn);

                if (string.IsNullOrWhiteSpace(connectionString))
                {
                    throw new InvalidOperationException("Connection has no connection string configured");
                }

                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref conn);
            }
        });
    }
}
