using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// VBA module and procedure operations for macro-enabled workbooks (.xlsm).
///
/// PREREQUISITES:
/// - Workbook must be macro-enabled (.xlsm)
/// - VBA trust must be enabled manually in Excel for project inspection and editing
///
/// SCOPE:
/// - List and view existing VBA components and their procedures
/// - Import creates new standard modules from inline code or a file
/// - Update/delete works on existing VBA components by name
/// - Run executes a procedure by name
///
/// RUN: procedureName format is 'Module.Procedure' (e.g., 'Module1.MySub').
/// Running an existing macro does not require VBA project access.
/// ExcelMcp does not configure VBA trust settings for you.
/// </summary>
[ServiceCategory("Vba")]
[McpTool("vba", Title = "VBA Operations", Destructive = true, Category = "automation",
    Description = "Inspect VBA project status and library references, search source with limited results, list procedures and source ranges, read a bounded part of a module, replace one procedure only when its source fingerprint still matches, import or update modules, and run procedures in .xlsm workbooks. Project inspection requires Trust Center access; status reports blocked access without enabling it. Replacing source preserves surrounding comments but does not prove it compiles or runs correctly.")]
[McpReadOnlyActions("list", "view", "read", "search", "references", "status")]
public interface IVbaCommands
{
    /// <summary>
    /// Lists all VBA modules and procedures in the workbook
    /// </summary>
    [ServiceAction("list")]
    VbaListResult List(IExcelBatch batch);

    /// <summary>
    /// Reports actual VBA project access, password protection, and execution mode without changing settings.
    /// </summary>
    [ServiceAction("status")]
    VbaProjectStatusResult Status(IExcelBatch batch);

    /// <summary>
    /// Lists VBA library references and flags missing libraries without repairing or changing references.
    /// </summary>
    [ServiceAction("references")]
    VbaReferencesResult References(IExcelBatch batch);

    /// <summary>
    /// Searches VBA source for literal text, returning limited matches with line numbers and excerpts.
    /// </summary>
    /// <param name="searchText">Nonempty, single-line literal text to find; wildcard patterns are not supported</param>
    /// <param name="moduleName">Optional module to search; omitted searches all modules in this workbook</param>
    /// <param name="wholeWord">Match whole words only; default false</param>
    /// <param name="matchCase">Match letter case; default false</param>
    /// <param name="maxMatches">Maximum returned matches, from 1 through 100; default 50. HasMore indicates omitted matches. Excerpts are limited to 200 characters.</param>
    [ServiceAction("search")]
    VbaSearchResult Search(
        IExcelBatch batch,
        [RequiredParameter] string searchText,
        string? moduleName = null,
        bool wholeWord = false,
        bool matchCase = false,
        int maxMatches = 50);

    /// <summary>
    /// Views VBA module code without exporting to file
    /// </summary>
    /// <param name="moduleName">Name of the VBA module</param>
    [ServiceAction("view")]
    VbaViewResult View(IExcelBatch batch, [RequiredParameter] string moduleName);

    /// <summary>
    /// Reads one VBA procedure or a bounded range of module lines.
    /// </summary>
    /// <param name="moduleName">Name of the VBA module</param>
    /// <param name="procedureName">Procedure to read; select this or startLine and lineCount</param>
    /// <param name="procedureKind">Optional kind to distinguish property accessors or same-name procedures</param>
    /// <param name="startLine">First module line to read; select this with lineCount instead of procedureName</param>
    /// <param name="lineCount">Number of module lines to read, from 1 through 500</param>
    [ServiceAction("read")]
    VbaReadResult Read(
        IExcelBatch batch,
        [RequiredParameter] string moduleName,
        string? procedureName,
        string? procedureKind,
        int? startLine,
        int? lineCount);

    /// <summary>
    /// Imports VBA code to create a new standard module
    /// </summary>
    /// <param name="moduleName">Name for the new module</param>
    /// <param name="vbaCode">VBA code. Public callers must supply either inline vbaCode or a readable vbaCodeFile, not both.</param>
    [ServiceAction("import")]
    OperationResult Import(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Updates an existing VBA module with new code
    /// </summary>
    /// <param name="moduleName">Name of the module to update</param>
    /// <param name="vbaCode">New VBA code. Public callers must supply either inline vbaCode or a readable vbaCodeFile, not both.</param>
    [ServiceAction("update")]
    OperationResult Update(IExcelBatch batch, [RequiredParameter] string moduleName, [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Replaces one VBA procedure only if its source has not changed since it was read, preserving surrounding comments and blank lines.
    /// </summary>
    /// <param name="moduleName">Name of the VBA module</param>
    /// <param name="procedureName">Name of the procedure to replace</param>
    /// <param name="procedureKind">Kind of procedure: Sub, Function, Property Get, Property Let, or Property Set</param>
    /// <param name="expectedSourceHash">SourceHash returned by vba.read for this procedure</param>
    /// <param name="vbaCode">Source for exactly one replacement procedure; saving it does not prove it compiles or runs</param>
    [ServiceAction("replace-procedure")]
    VbaProcedureEditResult ReplaceProcedure(
        IExcelBatch batch,
        [RequiredParameter] string moduleName,
        [RequiredParameter] string procedureName,
        [RequiredParameter] string procedureKind,
        [RequiredParameter] string expectedSourceHash,
        [RequiredParameter][FileOrValue] string vbaCode);

    /// <summary>
    /// Runs a VBA procedure with optional parameters
    /// </summary>
    /// <param name="procedureName">Name of the procedure to run (for example "Module1.MySub")</param>
    /// <param name="timeout">Optional public timeout in whole seconds from 1 through 2147483; converted to TimeSpan at shared dispatch</param>
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
