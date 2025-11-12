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
    /// Lists all connections in a workbook (filePath-based API)
    /// </summary>
    Task<ConnectionListResult> ListAsync(string filePath);

    /// <summary>
    /// Views detailed connection information
    /// </summary>
    Task<ConnectionViewResult> ViewAsync(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Views detailed connection information (filePath-based API)
    /// </summary>
    Task<ConnectionViewResult> ViewAsync(string filePath, string connectionName);

    /// <summary>
    /// Creates a new connection in the workbook
    /// </summary>
    Task<OperationResult> CreateAsync(IExcelBatch batch, string connectionName,
        string connectionString, string? commandText = null, string? description = null);

    /// <summary>
    /// Creates a new connection in the workbook (filePath-based API)
    /// </summary>
    Task<OperationResult> CreateAsync(string filePath, string connectionName,
        string connectionString, string? commandText = null, string? description = null);

    /// <summary>
    /// Imports connection from JSON file
    /// </summary>
    Task<OperationResult> ImportAsync(IExcelBatch batch, string connectionName, string jsonFilePath);

    /// <summary>
    /// Imports connection from JSON file (filePath-based API)
    /// </summary>
    Task<OperationResult> ImportAsync(string filePath, string connectionName, string jsonFilePath);

    /// <summary>
    /// Exports connection to JSON file
    /// </summary>
    Task<OperationResult> ExportAsync(IExcelBatch batch, string connectionName, string jsonFilePath);

    /// <summary>
    /// Exports connection to JSON file (filePath-based API)
    /// </summary>
    Task<OperationResult> ExportAsync(string filePath, string connectionName, string jsonFilePath);

    /// <summary>
    /// Updates existing connection properties from JSON file
    /// </summary>
    Task<OperationResult> UpdatePropertiesAsync(IExcelBatch batch, string connectionName, string jsonFilePath);

    /// <summary>
    /// Updates existing connection properties from JSON file (filePath-based API)
    /// </summary>
    Task<OperationResult> UpdatePropertiesAsync(string filePath, string connectionName, string jsonFilePath);

    /// <summary>
    /// Refreshes connection data
    /// </summary>
    Task<OperationResult> RefreshAsync(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Refreshes connection data (filePath-based API)
    /// </summary>
    Task<OperationResult> RefreshAsync(string filePath, string connectionName);

    /// <summary>
    /// Refreshes connection data with timeout
    /// </summary>
    Task<OperationResult> RefreshAsync(IExcelBatch batch, string connectionName, TimeSpan? timeout);

    /// <summary>
    /// Refreshes connection data with timeout (filePath-based API)
    /// </summary>
    Task<OperationResult> RefreshAsync(string filePath, string connectionName, TimeSpan? timeout);

    /// <summary>
    /// Deletes a connection
    /// </summary>
    Task<OperationResult> DeleteAsync(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Deletes a connection (filePath-based API)
    /// </summary>
    Task<OperationResult> DeleteAsync(string filePath, string connectionName);

    /// <summary>
    /// Loads connection data to a worksheet
    /// </summary>
    Task<OperationResult> LoadToAsync(IExcelBatch batch, string connectionName, string sheetName);

    /// <summary>
    /// Loads connection data to a worksheet (filePath-based API)
    /// </summary>
    Task<OperationResult> LoadToAsync(string filePath, string connectionName, string sheetName);

    /// <summary>
    /// Gets connection properties
    /// </summary>
    Task<ConnectionPropertiesResult> GetPropertiesAsync(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Gets connection properties (filePath-based API)
    /// </summary>
    Task<ConnectionPropertiesResult> GetPropertiesAsync(string filePath, string connectionName);

    /// <summary>
    /// Sets connection properties
    /// </summary>
    Task<OperationResult> SetPropertiesAsync(IExcelBatch batch, string connectionName,
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null,
        bool? savePassword = null, int? refreshPeriod = null);

    /// <summary>
    /// Sets connection properties (filePath-based API)
    /// </summary>
    Task<OperationResult> SetPropertiesAsync(string filePath, string connectionName,
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null,
        bool? savePassword = null, int? refreshPeriod = null);

    /// <summary>
    /// Tests connection without refreshing data
    /// </summary>
    Task<OperationResult> TestAsync(IExcelBatch batch, string connectionName);

    /// <summary>
    /// Tests connection without refreshing data (filePath-based API)
    /// </summary>
    Task<OperationResult> TestAsync(string filePath, string connectionName);
}
