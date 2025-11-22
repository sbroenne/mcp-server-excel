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
        string? connectionString = null, string? commandText = null, string? description = null,
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

            // Build connection definition with specified properties
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

            // Use UpdateConnectionProperties to apply all changes
            try
            {
                UpdateConnectionProperties(conn, definition);
            }
            catch (InvalidOperationException ex) when (ex.Message.Contains("0x800A03EC") && !string.IsNullOrWhiteSpace(connectionString))
            {
                // Excel blocks connection string updates for ODC-imported connections (security feature)
                throw new InvalidOperationException(
                    $"Cannot update connection string for connection '{connectionName}'. " +
                    "Excel blocks connection string changes for ODC-imported connections (security restriction). " +
                    "To change the data source, delete this connection and import a new ODC file, or create a new connection with excel_connection create action.",
                    ex);
            }

            result.Success = true;
            return result;
        });
    }
}

