using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Interface for connection management commands
/// </summary>
public interface IConnectionCommands
{
    /// <summary>
    /// Lists all connections in a workbook
    /// </summary>
    ConnectionListResult List(IExcelBatch batch);

    /// <summary>
    /// Views detailed connection information
    /// </summary>
    ConnectionViewResult View(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Creates a new connection in the workbook
    /// </summary>
    OperationResult Create(IExcelBatch batch, string connectionName,
        string connectionString, string? commandText = null, string? description = null);

    /// <summary>
    /// Imports connection from JSON file
    /// </summary>
    OperationResult Import(IExcelBatch batch, string connectionName, string jsonFilePath);

    /// <summary>
    /// Updates existing connection properties from JSON file
    /// </summary>
    OperationResult UpdateProperties(IExcelBatch batch, string connectionName, string jsonFilePath);

    /// <summary>
    /// Refreshes connection data
    /// </summary>
    OperationResult Refresh(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Refreshes connection data with timeout
    /// </summary>
    OperationResult Refresh(IExcelBatch batch, string connectionName, TimeSpan? timeout);

    /// <summary>
    /// Deletes a connection
    /// </summary>
    OperationResult Delete(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Loads connection data to a worksheet
    /// </summary>
    OperationResult LoadTo(IExcelBatch batch, string connectionName, string sheetName);

    /// <summary>
    /// Gets connection properties
    /// </summary>
    ConnectionPropertiesResult GetProperties(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Sets connection properties
    /// </summary>
    OperationResult SetProperties(IExcelBatch batch, string connectionName,
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null,
        bool? savePassword = null, int? refreshPeriod = null);

    /// <summary>
    /// Tests connection without refreshing data
    /// </summary>
    OperationResult Test(IExcelBatch batch, string connectionName);
}

