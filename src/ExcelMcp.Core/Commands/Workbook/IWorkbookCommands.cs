using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Workbook-level protection and view operations.
/// </summary>
[ServiceCategory("workbook", "Workbook")]
[McpTool("workbook", Title = "Workbook Operations", Destructive = true, Category = "structure",
    Description = "Workbook-level operations such as protecting or unprotecting the workbook and adjusting workbook view options.")]
public interface IWorkbookCommands
{
    /// <summary>
    /// Protects or unprotects the workbook.
    /// </summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="isProtected">Whether the workbook should be protected.</param>
    /// <param name="password">Optional password for protecting/unprotecting the workbook.</param>
    [ServiceAction("set-protection")]
    OperationResult SetProtection(
        IExcelBatch batch,
        [RequiredParameter] bool isProtected,
        string? password = null);

    /// <summary>
    /// Gets whether the workbook is protected.
    /// </summary>
    /// <param name="batch">Excel batch session.</param>
    [ServiceAction("get-protection")]
    WorkbookProtectionResult GetProtection(IExcelBatch batch);

    /// <summary>
    /// Sets workbook display options such as gridlines and headings.
    /// </summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="displayGridlines">Whether to display worksheet gridlines in the active workbook window.</param>
    /// <param name="displayHeadings">Whether to display row and column headings in the active workbook window.</param>
    [ServiceAction("set-view-options")]
    OperationResult SetViewOptions(
        IExcelBatch batch,
        bool? displayGridlines = null,
        bool? displayHeadings = null);

    /// <summary>
    /// Gets workbook display options such as gridlines and headings.
    /// </summary>
    /// <param name="batch">Excel batch session.</param>
    [ServiceAction("get-view-options")]
    WorkbookViewOptionsResult GetViewOptions(IExcelBatch batch);
}
