using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands.Workbook;

/// <summary>
/// Manage workbook metadata, integrity validation, document properties, Save As/copy operations, fixed-format exports, and external Excel links.
/// SAVE-AS formats: auto, xlsx, xlsm, xlsb, xls. The active session follows the new workbook path.
/// FIXED FORMAT: PDF or XPS with standard or minimum quality.
/// DOCUMENT PROPERTIES: built-in properties can be read/updated; custom properties can be created, updated, and deleted.
/// INTEGRITY: read-only checks for formula errors, external links, tables, and caller-supplied control totals.
/// EXTERNAL LINKS: discovers, updates, or permanently breaks Excel workbook links.
/// Printing and print preview are intentionally excluded because default-printer output and modal preview are unsafe for unattended automation.
/// </summary>
[ServiceCategory("workbook", "Workbook")]
[McpTool("workbook", Title = "Workbook Operations", Destructive = true, Category = "structure",
    Description = "Manage workbook metadata, read-only integrity validation, document properties, Save As/copy operations, fixed-format PDF/XPS exports, and external Excel links. INTEGRITY: validates formula results, external links, worksheet tables, and caller-supplied control totals without calculating, refreshing, editing, or saving. SAVE-AS formats: auto, xlsx, xlsm, xlsb, xls; the active session follows the new path. DOCUMENT PROPERTIES: built-in properties can be read/updated; custom properties can be created, updated, and deleted. EXTERNAL LINKS: list, update, or permanently break Excel workbook links. Printing and print preview are excluded because default-printer output and modal preview are unsafe for unattended automation.")]
public interface IWorkbookCommands
{
    /// <summary>Gets metadata for the active workbook.</summary>
    [ServiceAction("get-info")]
    WorkbookInfoResult GetInfo(IExcelBatch batch);

    /// <summary>Performs read-only workbook integrity checks without calculating, refreshing, editing, or saving.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="checks">Checks to run; omit for formulas, links, tables, and supplied control totals</param>
    /// <param name="worksheetNames">Optional worksheet names limiting formula and table checks</param>
    /// <param name="controlTotals">Expected numeric cells with optional absolute tolerances</param>
    /// <param name="maxFindings">Maximum finding details to return; counts still include omitted details. Range: 1-10000.</param>
    [ServiceAction("validate-integrity")]
    WorkbookIntegrityResult ValidateIntegrity(
        IExcelBatch batch,
        List<WorkbookIntegrityCheck>? checks = null,
        List<string>? worksheetNames = null,
        List<WorkbookControlTotalExpectation>? controlTotals = null,
        int maxFindings = 500);

    /// <summary>Lists built-in and/or custom workbook document properties.</summary>
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
    [ServiceAction("update-external-link")]
    OperationResult UpdateExternalLink(IExcelBatch batch, [RequiredParameter] string linkSource);

    /// <summary>Permanently breaks one external Excel workbook link, replacing formulas with their current values.</summary>
    [ServiceAction("break-external-link")]
    OperationResult BreakExternalLink(IExcelBatch batch, [RequiredParameter] string linkSource);

    /// <summary>Protects or unprotects the workbook structure.</summary>
    [ServiceAction("set-protection")]
    OperationResult SetProtection(
        IExcelBatch batch,
        [RequiredParameter] bool isProtected,
        string? password = null);

    /// <summary>Gets whether the workbook structure or windows are protected.</summary>
    [ServiceAction("get-protection")]
    WorkbookProtectionResult GetProtection(IExcelBatch batch);

    /// <summary>Sets workbook display options such as gridlines and headings.</summary>
    [ServiceAction("set-view-options")]
    OperationResult SetViewOptions(
        IExcelBatch batch,
        bool? displayGridlines = null,
        bool? displayHeadings = null);

    /// <summary>Gets workbook display options such as gridlines and headings.</summary>
    [ServiceAction("get-view-options")]
    WorkbookViewOptionsResult GetViewOptions(IExcelBatch batch);
}
