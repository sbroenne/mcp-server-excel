using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Interface for connection management commands
/// </summary>
[ServiceCategory("connection", "Connection")]
[McpTool("excel_connection")]
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
    [ServiceAction("view")]
    ConnectionViewResult View(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Creates a new connection in the workbook
    /// </summary>
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
    [ServiceAction("refresh")]
    void Refresh(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        [FromString("timeout")] TimeSpan? timeout = null);

    /// <summary>
    /// Deletes a connection
    /// </summary>
    [ServiceAction("delete")]
    void Delete(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Loads connection data to a worksheet
    /// </summary>
    [ServiceAction("load-to")]
    void LoadTo(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName,
        [RequiredParameter, FromString("sheetName")] string sheetName);

    /// <summary>
    /// Gets connection properties
    /// </summary>
    [ServiceAction("get-properties")]
    ConnectionPropertiesResult GetProperties(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);

    /// <summary>
    /// Sets connection properties (connection string, command text, description, and behavior settings)
    /// </summary>
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
    [ServiceAction("test")]
    void Test(
        IExcelBatch batch,
        [RequiredParameter, FromString("connectionName")] string connectionName);
}



