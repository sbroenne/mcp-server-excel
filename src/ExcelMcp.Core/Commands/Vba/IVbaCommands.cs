using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA module and procedure operations for macro-enabled workbooks (.xlsm).
///
/// PREREQUISITES:
/// - Workbook must be macro-enabled (.xlsm)
/// - VBA trust must be enabled manually in Excel for project access
///
/// SCOPE:
/// - List and view existing VBA components and their procedures
/// - Import creates new standard modules from inline code or a file
/// - Update/delete works on existing VBA components by name
/// - Run executes a procedure by name
///
/// RUN: procedureName format is 'Module.Procedure' (e.g., 'Module1.MySub').
/// ExcelMcp does not configure VBA trust settings for you.
/// </summary>
[ServiceCategory("vba", "Vba")]
[McpTool("vba", Title = "VBA Operations", Destructive = true, Category = "automation",
    Description = "VBA module and procedure operations for macro-enabled workbooks (.xlsm). Lists and views existing VBA components, imports new standard modules, updates or deletes module code, and runs procedures. VBA trust must be enabled manually in Excel; ExcelMcp does not configure Trust Center settings.")]
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
    /// <param name="moduleName">Name of the VBA module</param>
    [ServiceAction("view")]
    VbaViewResult View(IExcelBatch batch, [RequiredParameter] string moduleName);

    /// <summary>
    /// Imports VBA code to create a new standard module
    /// </summary>
    /// <param name="moduleName">Name for the new module</param>
    /// <param name="vbaCode">VBA code to import</param>
    [ServiceAction("import")]
    OperationResult Import(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    /// <param name="moduleName">Name of the module to update</param>
    /// <param name="vbaCode">New VBA code</param>
    [ServiceAction("update")]
    OperationResult Update(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    /// <param name="procedureName">Name of the procedure to run (for example "Module1.MySub")</param>
    /// <param name="timeout">Optional timeout for execution</param>
    /// <param name="parameters">Optional parameters to pass to the procedure</param>
    [ServiceAction("run")]
    OperationResult Run(IExcelBatch batch, [RequiredParameter] string procedureName, TimeSpan? timeout, params string[] parameters);

    /// <summary>
    /// Deletes a VBA module
    /// </summary>
    /// <param name="moduleName">Name of the module to delete</param>
    [ServiceAction("delete")]
    OperationResult Delete(IExcelBatch batch, [RequiredParameter] string moduleName);
}



