using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA script management commands
/// </summary>
[ServiceCategory("vba", "Vba")]
[McpTool("excel_vba")]
public interface IVbaCommands
{
    /// <summary>
    /// Lists all VBA modules and procedures in the workbook
    /// </summary>
    [ServiceAction("list")]
    VbaListResult List(IExcelBatch batch);

    /// <summary>
    /// Views VBA module code without exporting to file
    /// </summary>
    [ServiceAction("view")]
    VbaViewResult View(IExcelBatch batch, [RequiredParameter] string moduleName);

    /// <summary>
    /// Imports VBA code to create a new module
    /// </summary>
    [ServiceAction("import")]
    void Import(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    [ServiceAction("update")]
    void Update(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    [ServiceAction("run")]
    void Run(IExcelBatch batch, [RequiredParameter] string procedureName, TimeSpan? timeout, params string[] parameters);

    /// <summary>
    /// Deletes a VBA module
    /// </summary>
    [ServiceAction("delete")]
    void Delete(IExcelBatch batch, [RequiredParameter] string moduleName);
}



