using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
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
            dynamic? sheets = null;
            dynamic? targetSheet = null;

            try
            {
                conn = ComUtilities.FindConnection(ctx.Book, connectionName);

                if (conn == null)
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' not found.");
                }

                // Check if this is a Power Query connection
                if (PowerQueryHelpers.IsPowerQueryConnection(conn))
                {
                    throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Use 'pq-loadto' command instead.");
                }

                // Find or create target sheet
                sheets = ctx.Book.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    dynamic? sheet = null;
                    try
                    {
                        sheet = sheets.Item(i);
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
                    targetSheet = sheets.Add();
                    targetSheet.Name = sheetName;
                }

                // Remove existing QueryTables first
                PowerQueryHelpers.RemoveQueryTables(ctx.Book, connectionName);

                // Create QueryTable to load data
                var options = new PowerQueryHelpers.QueryTableOptions
                {
                    Name = connectionName,
                    RefreshImmediately = true
                };

                CreateQueryTableForConnection(targetSheet, conn, options);
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }
            finally
            {
                ComUtilities.Release(ref targetSheet);
                ComUtilities.Release(ref sheets);
                ComUtilities.Release(ref conn);
            }
        }, timeoutCts.Token);
    }

    /// <summary>
    /// Gets connection properties
    /// </summary>

    public OperationResult Test(IExcelBatch batch, string connectionName)
    {
        return batch.Execute((ctx, ct) =>
        {
            Excel.WorkbookConnection? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                throw new InvalidOperationException($"Connection '{connectionName}' not found.");
            }

            // Get connection type
            int connType = (int)conn.Type;

            // For Text (4) and Web (5) connections, connection string might not be accessible
            // until a QueryTable is created. Just verify the connection object exists.
            if (connType is 4 or 5)
            {
                return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
            }

            // For other connection types (OLEDB, ODBC), validate connection string
            string? connectionString = GetConnectionString(conn);

            if (string.IsNullOrWhiteSpace(connectionString))
            {
                throw new InvalidOperationException("Connection has no connection string configured");
            }

            // Connection exists and is accessible
            return new OperationResult { Success = true, FilePath = batch.WorkbookPath };
        });
    }
}



