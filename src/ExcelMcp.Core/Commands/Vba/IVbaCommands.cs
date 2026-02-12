using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA scripts (requires .xlsm and VBA trust enabled).
///
/// PREREQUISITES:
/// - Workbook must be macro-enabled (.xlsm)
/// - VBA trust must be enabled for automation
///
/// RUN: procedureName format is 'Module.Procedure' (e.g., 'Module1.MySub').
/// </summary>
[ServiceCategory("vba", "Vba")]
[McpTool("vba", Title = "VBA Operations", Destructive = true, Category = "automation",
    Description = "VBA scripts (requires .xlsm and VBA trust enabled). Manages VBA macro operations, code import/export, and script execution in macro-enabled workbooks. Prerequisites: Use setup-vba-trust to configure VBA trust for automation.")]
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
    /// Imports VBA code to create a new module
    /// </summary>
    /// <param name="moduleName">Name for the new module</param>
    /// <param name="vbaCode">VBA code to import</param>
    [ServiceAction("import")]
    void Import(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    /// <param name="moduleName">Name of the module to update</param>
    /// <param name="vbaCode">New VBA code</param>
    [ServiceAction("update")]
    void Update(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    /// <param name="procedureName">Name of the procedure to run (e.g., "Module1.MySub")</param>
    /// <param name="timeout">Optional timeout for execution</param>
    /// <param name="parameters">Optional parameters to pass to the procedure</param>
    [ServiceAction("run")]
    void Run(IExcelBatch batch, [RequiredParameter] string procedureName, TimeSpan? timeout, params string[] parameters);

    /// <summary>
    /// Deletes a VBA module
    /// </summary>
    /// <param name="moduleName">Name of the module to delete</param>
    [ServiceAction("delete")]
    void Delete(IExcelBatch batch, [RequiredParameter] string moduleName);
}



