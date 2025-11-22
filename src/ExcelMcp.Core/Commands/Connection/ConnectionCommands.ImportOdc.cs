using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Connection import operations (ImportFromOdc)
/// </summary>
public partial class ConnectionCommands
{
    /// <summary>
    /// Imports a connection from an ODC file
    /// </summary>
    public OperationResult ImportFromOdc(IExcelBatch batch, string odcFilePath)
    {
        var result = new OperationResult
        {
            FilePath = batch.WorkbookPath,
            Action = "import-odc"
        };

        // Validate file exists before attempting import
        if (!File.Exists(odcFilePath))
        {
            result.Success = false;
            result.ErrorMessage = $"ODC file not found: {odcFilePath}";
            return result;
        }

        return batch.Execute((ctx, ct) =>
        {
            dynamic? connections = null;
            dynamic? conn = null;

            try
            {
                connections = ctx.Book.Connections;
                int countBefore = connections.Count;

                // Import ODC file via Excel COM API
                // Excel handles all ODC parsing and validation
                conn = connections.AddFromFile(odcFilePath);

                // Verify connection was imported
                int countAfter = connections.Count;
                if (countAfter > countBefore && conn != null)
                {
                    string connectionName = conn.Name?.ToString() ?? "Unknown";

                    result.Success = true;
                    result.Action = $"import-odc (imported: {connectionName})";
                }
                else
                {
                    result.Success = false;
                    result.ErrorMessage = "ODC file imported but no new connection was created";
                }

                return result;
            }
            finally
            {
                ComUtilities.Release(ref conn);
                ComUtilities.Release(ref connections);
            }
        });
    }
}
