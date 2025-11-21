using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;

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
            dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                throw new InvalidOperationException($"Connection '{connectionName}' not found");
            }

            result.BackgroundQuery = GetBackgroundQuerySetting(conn);
            result.RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn);
            result.SavePassword = GetSavePasswordSetting(conn);
            result.RefreshPeriod = GetRefreshPeriod(conn);

            result.Success = true;
            return result;
        });
    }

    /// <inheritdoc />
    public OperationResult SetProperties(IExcelBatch batch, string connectionName,
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null,
        bool? savePassword = null, int? refreshPeriod = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-properties"
        };

        return batch.Execute((ctx, ct) =>
        {
            dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

            if (conn == null)
            {
                throw new InvalidOperationException($"Connection '{connectionName}' not found");
            }

            // Check if this is a Power Query connection
            if (PowerQueryHelpers.IsPowerQueryConnection(conn))
            {
                throw new InvalidOperationException($"Connection '{connectionName}' is a Power Query connection. Power Query properties cannot be modified directly.");
            }

            // Update properties if specified
            SetConnectionProperty(conn, "BackgroundQuery", backgroundQuery);
            SetConnectionProperty(conn, "RefreshOnFileOpen", refreshOnFileOpen);
            SetConnectionProperty(conn, "SavePassword", savePassword);
            SetConnectionProperty(conn, "RefreshPeriod", refreshPeriod);

            result.Success = true;
            return result;
        });
    }
}

