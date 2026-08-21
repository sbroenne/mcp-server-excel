using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data connections (OLEDB, ODBC, ODC import).
/// TEXT/WEB/CSV: Use querytable for direct local imports or powerquery for transformations.
/// Power Query connections auto-redirect to powerquery.
/// TIMEOUT: 30 min auto-timeout for refresh/load-to.
/// </summary>
[ServiceCategory("connection", "Connection")]
[McpTool("connection", Title = "Data Connection Operations", Destructive = true, Category = "query",
    Description = "Data connections (OLEDB, ODBC, ODC import). TEXT/WEB/CSV: Use querytable for direct local imports or powerquery for transformations. Power Query connections redirect to powerquery by exact mashup Location identity. Delete/load-to cleanup follows the exact WorkbookConnection and preserves unrelated similarly named QueryTables. Typed OLEDB/ODBC refresh status and cancellation are available. TIMEOUT: 30 min auto-timeout for refresh/loadto.")]
public interface IConnectionCommands
{
    /// <summary>
    /// Lists all connections in a workbook
    /// </summary>
    [ServiceAction("list")]
    ConnectionListResult List(IExcelBatch batch);

    /// <summary>
    /// Views detailed connection information
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection to view</param>
    [ServiceAction("view")]
    ConnectionViewResult View(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Creates a new connection in the workbook
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name for the new connection</param>
    /// <param name="connectionString">OLEDB or ODBC connection string</param>
    /// <param name="commandText">SQL query or table name</param>
    /// <param name="description">Optional description for the connection</param>
    [ServiceAction("create")]
    OperationResult Create(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        [RequiredParameter, FromString("connectionString")] string connectionString,
        [FromString("commandText")] string? commandText = null,
        [FromString("description")] string? description = null);

    /// <summary>
    /// Refreshes connection data with optional timeout
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection to refresh</param>
    /// <param name="timeout">Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch</param>
    [ServiceAction("refresh")]
    OperationResult Refresh(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        [FromString("timeout")] TimeSpan? timeout = null);

    /// <summary>
    /// Gets refresh status for OLEDB and ODBC background refreshes started
    /// outside the synchronous connection refresh action.
    /// Excel PIA exposes status on the typed sub-connection, not WorkbookConnection.
    /// The synchronous refresh action occupies the session queue until completion.
    /// </summary>
    [ServiceAction("get-refresh-status")]
    RefreshStatusResult GetRefreshStatus(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Cancels an active OLEDB or ODBC background refresh started outside the
    /// synchronous connection refresh action.
    /// Returns an explicit unsupported result for connection types without a typed PIA cancellation API.
    /// </summary>
    [ServiceAction("cancel-refresh")]
    RefreshCancellationResult CancelRefresh(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Deletes a connection and QueryTables owned by that exact WorkbookConnection.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection to delete</param>
    [ServiceAction("delete")]
    OperationResult Delete(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Loads connection data to a worksheet, replacing only QueryTables owned by
    /// the exact WorkbookConnection.
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection</param>
    /// <param name="sheetName">Target worksheet name</param>
    [ServiceAction("load-to")]
    OperationResult LoadTo(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        [RequiredParameter, FromString("sheetName")] string sheetName);

    /// <summary>
    /// Gets connection properties
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection</param>
    [ServiceAction("get-properties")]
    ConnectionPropertiesResult GetProperties(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Sets connection properties (connection string, command text, description, and behavior settings)
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection</param>
    /// <param name="connectionString">New connection string (null to keep current)</param>
    /// <param name="commandText">New SQL query or table name (null to keep current)</param>
    /// <param name="description">New description (null to keep current)</param>
    /// <param name="backgroundQuery">Run query in background (null to keep current)</param>
    /// <param name="refreshOnFileOpen">Refresh when file opens (null to keep current)</param>
    /// <param name="savePassword">Save password in connection (null to keep current)</param>
    /// <param name="refreshPeriod">Auto-refresh interval in minutes (null to keep current)</param>
    [ServiceAction("set-properties")]
    OperationResult SetProperties(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        string? connectionString = null,
        string? commandText = null,
        string? description = null,
        bool? backgroundQuery = null,
        bool? refreshOnFileOpen = null,
        bool? savePassword = null,
        int? refreshPeriod = null);

    /// <summary>
    /// Tests connection without refreshing data
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection to test</param>
    [ServiceAction("test")]
    OperationResult Test(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);
}
