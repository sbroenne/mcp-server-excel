using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

public partial class ConnectionCommands
{
    /// <inheritdoc />
    public RefreshStatusResult GetRefreshStatus(IExcelBatch batch, string connectionName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? connection = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName);
            try
            {
                if (connection == null)
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                }

                var (supported, refreshing) = ReadRefreshStatus(connection);
                return new RefreshStatusResult
                {
                    Success = true,
                    FilePath = batch.WorkbookPath,
                    SupportsRefreshStatus = supported,
                    IsRefreshing = refreshing
                };
            }
            finally
            {
                ComUtilities.Release(ref connection);
            }
        });
    }

    /// <inheritdoc />
    public RefreshCancellationResult CancelRefresh(IExcelBatch batch, string connectionName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? connection = PowerQueryHelpers.FindConnectionByExactName(ctx.Book, connectionName);
            try
            {
                if (connection == null)
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                }

                return CancelTypedRefresh(connection, batch.WorkbookPath);
            }
            finally
            {
                ComUtilities.Release(ref connection);
            }
        });
    }

    private static (bool Supported, bool Refreshing) ReadRefreshStatus(Excel.WorkbookConnection connection)
    {
        if (connection.Type == Excel.XlConnectionType.xlConnectionTypeOLEDB)
        {
            Excel.OLEDBConnection? oledb = null;
            try
            {
                oledb = connection.OLEDBConnection;
                return (true, oledb.Refreshing);
            }
            finally
            {
                ComUtilities.Release(ref oledb);
            }
        }

        if (connection.Type == Excel.XlConnectionType.xlConnectionTypeODBC)
        {
            Excel.ODBCConnection? odbc = null;
            try
            {
                odbc = connection.ODBCConnection;
                return (true, odbc.Refreshing);
            }
            finally
            {
                ComUtilities.Release(ref odbc);
            }
        }

        return (false, false);
    }

    private static RefreshCancellationResult CancelTypedRefresh(
        Excel.WorkbookConnection connection,
        string workbookPath)
    {
        if (connection.Type == Excel.XlConnectionType.xlConnectionTypeOLEDB)
        {
            Excel.OLEDBConnection? oledb = null;
            try
            {
                oledb = connection.OLEDBConnection;
                bool wasRefreshing = oledb.Refreshing;
                if (wasRefreshing)
                {
                    oledb.CancelRefresh();
                }

                return CreateCancellationResult(workbookPath, true, wasRefreshing);
            }
            finally
            {
                ComUtilities.Release(ref oledb);
            }
        }

        if (connection.Type == Excel.XlConnectionType.xlConnectionTypeODBC)
        {
            Excel.ODBCConnection? odbc = null;
            try
            {
                odbc = connection.ODBCConnection;
                bool wasRefreshing = odbc.Refreshing;
                if (wasRefreshing)
                {
                    odbc.CancelRefresh();
                }

                return CreateCancellationResult(workbookPath, true, wasRefreshing);
            }
            finally
            {
                ComUtilities.Release(ref odbc);
            }
        }

        return CreateCancellationResult(workbookPath, false, false);
    }

    private static RefreshCancellationResult CreateCancellationResult(
        string workbookPath,
        bool supported,
        bool wasRefreshing)
    {
        return new RefreshCancellationResult
        {
            Success = true,
            FilePath = workbookPath,
            SupportsCancellation = supported,
            WasRefreshing = wasRefreshing,
            Cancelled = supported && wasRefreshing
        };
    }
}
