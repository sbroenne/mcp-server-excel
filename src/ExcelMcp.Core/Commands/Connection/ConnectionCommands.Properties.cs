using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;
using Excel = Microsoft.Office.Interop.Excel;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection property management (Get/Set properties)
/// </summary>
public partial class ConnectionCommands
{
    /// <inheritdoc />
    public ConnectionPropertiesResult GetProperties(IExcelBatch batch, string connectionName)
    {
        var result = new ConnectionPropertiesResult
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

                result.BackgroundQuery = GetBackgroundQuerySetting(conn);
                result.RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn);
                result.SavePassword = GetSavePasswordSetting(conn);
                result.RefreshPeriod = GetRefreshPeriod(conn);

                result.Success = true;
                return result;
            }
            finally
            {
                ComUtilities.Release(ref conn);
            }
        });
    }

    /// <inheritdoc />
    public OperationResult SetProperties(IExcelBatch batch, string connectionName,
        string? connectionString = null, string? commandText = null, string? description = null,
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null,
        bool? savePassword = null, int? refreshPeriod = null)
    {
        if (refreshPeriod is < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(refreshPeriod), refreshPeriod,
                "Refresh period must be nonnegative; use 0 to disable automatic refresh.");
        }

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
                    throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Power Query properties cannot be modified directly.");
                }

                var definition = new ConnectionDefinition
                {
                    ConnectionString = connectionString,
                    CommandText = commandText,
                    Description = description,
                    BackgroundQuery = backgroundQuery,
                    RefreshOnFileOpen = refreshOnFileOpen,
                    SavePassword = savePassword,
                    RefreshPeriod = refreshPeriod
                };

                try
                {
                    UpdateConnectionProperties(conn, definition);
                }
                catch (InvalidOperationException ex) when (ex.Message.Contains("0x800A03EC") && !string.IsNullOrWhiteSpace(connectionString))
                {
                    throw new InvalidOperationException(
                        $"Cannot update connection string for connection '{connectionName}'. " +
                        "Excel blocks connection string changes for ODC-imported connections (security restriction). " +
                        "To change the data source, delete this connection and import a new ODC file, or create a new connection with connection create action.",
                        ex);
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
