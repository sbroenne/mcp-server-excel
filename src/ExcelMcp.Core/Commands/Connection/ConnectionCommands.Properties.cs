using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Connections;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.PowerQuery;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection property management (Get/Set properties)
/// </summary>
public partial class ConnectionCommands
{
    /// <inheritdoc />
    public async Task<ConnectionPropertiesResult> GetPropertiesAsync(IExcelBatch batch, string connectionName)
    {
        var result = new ConnectionPropertiesResult
        {
            FilePath = batch.WorkbookPath,
            ConnectionName = connectionName
        };

        return await batch.Execute((ctx, ct) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

                if (conn == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' not found";
                    return result;
                }

                result.BackgroundQuery = GetBackgroundQuerySetting(conn);
                result.RefreshOnFileOpen = GetRefreshOnFileOpenSetting(conn);
                result.SavePassword = GetSavePasswordSetting(conn);
                result.RefreshPeriod = GetRefreshPeriod(conn);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error getting connection properties: {ex.Message}";
                return result;
            }
        });
    }

    /// <inheritdoc />
    public async Task<OperationResult> SetPropertiesAsync(IExcelBatch batch, string connectionName,
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null,
        bool? savePassword = null, int? refreshPeriod = null)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "set-properties"
        };

        return await batch.Execute((ctx, ct) =>
        {
            try
            {
                dynamic? conn = ComUtilities.FindConnection(ctx.Book, connectionName);

                if (conn == null)
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' not found";
                    return result;
                }

                // Check if this is a Power Query connection
                if (PowerQueryHelpers.IsPowerQueryConnection(conn))
                {
                    result.Success = false;
                    result.ErrorMessage = $"Connection '{connectionName}' is a Power Query connection. Power Query properties cannot be modified directly.";
                    return result;
                }

                // Update properties if specified
                SetConnectionProperty(conn, "BackgroundQuery", backgroundQuery);
                SetConnectionProperty(conn, "RefreshOnFileOpen", refreshOnFileOpen);
                SetConnectionProperty(conn, "SavePassword", savePassword);
                SetConnectionProperty(conn, "RefreshPeriod", refreshPeriod);

                result.Success = true;
                return result;
            }
            catch (Exception ex)
            {
                result.Success = false;
                result.ErrorMessage = $"Error setting connection properties: {ex.Message}";
                return result;
            }
        });
    }
}
