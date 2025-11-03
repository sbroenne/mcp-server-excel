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
    Task<ConnectionListResult> ListAsync(IExcelBatch batch);

    /// <summary>
    /// Views detailed connection information
    /// </summary>
    Task<ConnectionViewResult> ViewAsync(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Imports connection from JSON file
    /// </summary>
    Task<OperationResult> ImportAsync(IExcelBatch batch, string connectionName, string jsonFilePath);

    /// <summary>
    /// Exports connection to JSON file
    /// </summary>
    Task<OperationResult> ExportAsync(IExcelBatch batch, string connectionName, string jsonFilePath);

    /// <summary>
    /// Updates existing connection from JSON file
    /// </summary>
    Task<OperationResult> UpdateAsync(IExcelBatch batch, string connectionName, string jsonFilePath);

    /// <summary>
    /// Refreshes connection data
    /// </summary>
    Task<OperationResult> RefreshAsync(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Deletes a connection
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Loads connection data to a worksheet
    /// </summary>
    Task<OperationResult> LoadToAsync(IExcelBatch batch, string connectionName, string sheetName);

    /// <summary>
    /// Gets connection properties
    /// </summary>
    Task<ConnectionPropertiesResult> GetPropertiesAsync(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Sets connection properties
    /// </summary>
    Task<OperationResult> SetPropertiesAsync(IExcelBatch batch, string connectionName,
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null,
        bool? savePassword = null, int? refreshPeriod = null);

    /// <summary>
    /// Tests connection without refreshing data
    /// </summary>
    Task<OperationResult> TestAsync(IExcelBatch batch, string connectionName);
}
