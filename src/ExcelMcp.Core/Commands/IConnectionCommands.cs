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
    ConnectionListResult List(string filePath);

    /// <summary>
    /// Views detailed connection information
    /// </summary>
    ConnectionViewResult View(string filePath, string connectionName);

    /// <summary>
    /// Imports connection from JSON file
    /// </summary>
    OperationResult Import(string filePath, string connectionName, string jsonFilePath);

    /// <summary>
    /// Exports connection to JSON file
    /// </summary>
    OperationResult Export(string filePath, string connectionName, string jsonFilePath);

    /// <summary>
    /// Updates existing connection from JSON file
    /// </summary>
    OperationResult Update(string filePath, string connectionName, string jsonFilePath);

    /// <summary>
    /// Refreshes connection data
    /// </summary>
    OperationResult Refresh(string filePath, string connectionName);

    /// <summary>
    /// Deletes a connection
    /// </summary>
    OperationResult Delete(string filePath, string connectionName);

    /// <summary>
    /// Loads connection data to a worksheet
    /// </summary>
    OperationResult LoadTo(string filePath, string connectionName, string sheetName);

    /// <summary>
    /// Gets connection properties
    /// </summary>
    ConnectionPropertiesResult GetProperties(string filePath, string connectionName);

    /// <summary>
    /// Sets connection properties
    /// </summary>
    OperationResult SetProperties(string filePath, string connectionName, 
        bool? backgroundQuery = null, bool? refreshOnFileOpen = null, 
        bool? savePassword = null, int? refreshPeriod = null);

    /// <summary>
    /// Tests connection without refreshing data
    /// </summary>
    OperationResult Test(string filePath, string connectionName);
}
