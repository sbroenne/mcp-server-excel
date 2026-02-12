using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data connections (OLEDB, ODBC, ODC import).
/// TEXT/WEB/CSV: Use powerquery instead.
/// Power Query connections auto-redirect to powerquery.
/// TIMEOUT: 5 min auto-timeout for refresh/load-to.
/// </summary>
[ServiceCategory("connection", "Connection")]
[McpTool("connection", Title = "Data Connection Operations", Destructive = true, Category = "query",
    Description = "Data connections (OLEDB, ODBC, ODC import). TEXT/WEB/CSV: Use powerquery instead. Power Query connections auto-redirect to powerquery. TIMEOUT: 5 min auto-timeout for refresh/loadto.")]
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
    void Create(
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
    /// <param name="timeout">Optional timeout for the refresh operation</param>
    [ServiceAction("refresh")]
    void Refresh(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        [FromString("timeout")] TimeSpan? timeout = null);

    /// <summary>
    /// Deletes a connection
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection to delete</param>
    [ServiceAction("delete")]
    void Delete(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Loads connection data to a worksheet
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="connectionName">Name of the connection</param>
    /// <param name="sheetName">Target worksheet name</param>
    [ServiceAction("load-to")]
    void LoadTo(
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
    void SetProperties(
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
    void Test(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);
}



