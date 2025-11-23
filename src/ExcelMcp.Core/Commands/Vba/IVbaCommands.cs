using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands
/// </summary>
public interface IVbaCommands
{
    /// <summary>
    /// Lists all VBA modules and procedures in the workbook
    /// </summary>
    VbaListResult List(IExcelBatch batch);

    /// <summary>
    /// Views VBA module code without exporting to file
    /// </summary>
    VbaViewResult View(IExcelBatch batch, string moduleName);

    /// <summary>
    /// Imports VBA code to create a new module
    /// </summary>
    void Import(IExcelBatch batch, string moduleName, string vbaCode);

    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    void Update(IExcelBatch batch, string moduleName, string vbaCode);

    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    void Run(IExcelBatch batch, string procedureName, TimeSpan? timeout, params string[] parameters);

    /// <summary>
    /// Deletes a VBA module
    /// </summary>
    void Delete(IExcelBatch batch, string moduleName);
}

