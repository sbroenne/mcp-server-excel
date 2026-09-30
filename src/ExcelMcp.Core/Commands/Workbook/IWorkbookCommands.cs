using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// Manage workbook metadata, document properties, Save As/copy operations, fixed-format exports, and external Excel links.
/// SAVE-AS formats: auto, xlsx, xlsm, xlsb, xls. The active session follows the new workbook path.
/// FIXED FORMAT: PDF or XPS with standard or minimum quality.
/// DOCUMENT PROPERTIES: built-in properties can be read/updated; custom properties can be created, updated, and deleted.
/// EXTERNAL LINKS: discovers, updates, or permanently breaks Excel workbook links.
/// Printing and print preview are intentionally excluded because default-printer output and modal preview are unsafe for unattended automation.
/// </summary>
[ServiceCategory("workbook", "Workbook")]
[McpTool("workbook", Title = "Workbook Operations", Destructive = true, Category = "structure",
    Description = "Manage workbook metadata, document properties, Save As/copy operations, fixed-format PDF/XPS exports, and external Excel links. SAVE-AS formats: auto, xlsx, xlsm, xlsb, xls; the active session follows the new path. DOCUMENT PROPERTIES: built-in properties can be read/updated; custom properties can be created, updated, and deleted. EXTERNAL LINKS: list, update, or permanently break Excel workbook links. Printing and print preview are excluded because default-printer output and modal preview are unsafe for unattended automation.")]
public interface IWorkbookCommands
{
    /// <summary>Gets metadata for the active workbook.</summary>
    [ServiceAction("get-info")]
    WorkbookInfoResult GetInfo(IExcelBatch batch);

    /// <summary>Lists built-in and/or custom workbook document properties.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="includeBuiltIn">Include built-in document properties</param>
    /// <param name="includeCustom">Include custom document properties</param>
    [ServiceAction("list-document-properties")]
    DocumentPropertyListResult ListDocumentProperties(
        IExcelBatch batch,
        bool includeBuiltIn = true,
        bool includeCustom = true);

    /// <summary>Gets one built-in or custom workbook document property.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="propertyName">Document property name</param>
    /// <param name="scope">Property collection: built-in or custom</param>
    [ServiceAction("get-document-property")]
    DocumentPropertyResult GetDocumentProperty(
        IExcelBatch batch,
        [RequiredParameter] string propertyName,
        [FromString] DocumentPropertyScope scope = DocumentPropertyScope.Custom);

    /// <summary>Creates or updates a custom property, or updates an existing built-in property.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="propertyName">Document property name</param>
    /// <param name="value">String value to store</param>
    /// <param name="scope">Property collection: built-in or custom</param>
    [ServiceAction("set-document-property")]
    OperationResult SetDocumentProperty(
        IExcelBatch batch,
        [RequiredParameter] string propertyName,
        [RequiredParameter] string value,
        [FromString] DocumentPropertyScope scope = DocumentPropertyScope.Custom);

    /// <summary>Deletes a custom workbook document property. Built-in properties cannot be deleted.</summary>
    [ServiceAction("delete-document-property")]
    OperationResult DeleteDocumentProperty(IExcelBatch batch, [RequiredParameter] string propertyName);

    /// <summary>Saves the active workbook under a new path and format, then moves the active session to that path.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="targetPath">Absolute output path in an existing directory</param>
    /// <param name="format">Output format: auto, xlsx, xlsm, xlsb, or xls</param>
    /// <param name="overwrite">Whether an existing output file may be replaced</param>
    [ServiceAction("save-as")]
    OperationResult SaveAs(
        IExcelBatch batch,
        [RequiredParameter] string targetPath,
        [FromString] WorkbookSaveFormat format = WorkbookSaveFormat.Auto,
        bool overwrite = false);

    /// <summary>Saves a copy without changing the active workbook or session path. The output extension must match the active workbook.</summary>
    [ServiceAction("save-copy-as")]
    OperationResult SaveCopyAs(
        IExcelBatch batch,
        [RequiredParameter] string targetPath,
        bool overwrite = false);

    /// <summary>Exports the workbook to PDF or XPS using Excel's fixed-format renderer.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="targetPath">Absolute output path in an existing directory</param>
    /// <param name="formatType">Fixed-format output: Pdf or Xps</param>
    /// <param name="quality">Export quality: Standard or Minimum</param>
    /// <param name="includeDocumentProperties">Include document metadata in the exported file</param>
    /// <param name="ignorePrintAreas">Export without restricting output to configured print areas</param>
    /// <param name="fromPage">First page to export, 1-based; omit to start at the beginning</param>
    /// <param name="toPage">Last page to export, inclusive; omit to export through the end</param>
    /// <param name="openAfterPublish">Open the exported file in its associated viewer</param>
    /// <param name="overwrite">Whether an existing output file may be replaced</param>
    [ServiceAction("export-fixed-format")]
    OperationResult ExportFixedFormat(
        IExcelBatch batch,
        [RequiredParameter] string targetPath,
        [FromString] FixedFormatType formatType = FixedFormatType.Pdf,
        [FromString] FixedFormatQuality quality = FixedFormatQuality.Standard,
        bool includeDocumentProperties = true,
        bool ignorePrintAreas = false,
        int? fromPage = null,
        int? toPage = null,
        bool openAfterPublish = false,
        bool overwrite = false);

    /// <summary>Lists external Excel workbook links referenced by the active workbook.</summary>
    [ServiceAction("list-external-links")]
    ExternalLinkListResult ListExternalLinks(IExcelBatch batch);

    /// <summary>Updates one external Excel workbook link from its source.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="linkSource">Exact source identifier returned by list-external-links</param>
    [ServiceAction("update-external-link")]
    OperationResult UpdateExternalLink(IExcelBatch batch, [RequiredParameter] string linkSource);

    /// <summary>Permanently breaks one external Excel workbook link, replacing formulas with their current values.</summary>
    [ServiceAction("break-external-link")]
    OperationResult BreakExternalLink(IExcelBatch batch, [RequiredParameter] string linkSource);

    /// <summary>Protects or unprotects the workbook structure.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="isProtected">True to protect workbook structure, false to unprotect it</param>
    /// <param name="password">Optional protection password; required to unprotect password-protected structure</param>
    [ServiceAction("set-protection")]
    OperationResult SetProtection(
        IExcelBatch batch,
        [RequiredParameter] bool isProtected,
        string? password = null);

    /// <summary>Gets whether the workbook structure or windows are protected.</summary>
    [ServiceAction("get-protection")]
    WorkbookProtectionResult GetProtection(IExcelBatch batch);

    /// <summary>Sets workbook display options such as gridlines and headings.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="displayGridlines">Show or hide gridlines; omit to leave unchanged</param>
    /// <param name="displayHeadings">Show or hide row/column headings; omit to leave unchanged</param>
    [ServiceAction("set-view-options")]
    OperationResult SetViewOptions(
        IExcelBatch batch,
        bool? displayGridlines = null,
        bool? displayHeadings = null);

    /// <summary>Gets workbook display options such as gridlines and headings.</summary>
    [ServiceAction("get-view-options")]
    WorkbookViewOptionsResult GetViewOptions(IExcelBatch batch);
}
