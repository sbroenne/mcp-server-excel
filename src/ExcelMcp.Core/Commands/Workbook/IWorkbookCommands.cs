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
/// break-external-link replaces linked formulas with their current values; no tool-level undo.
/// Printing and print preview are intentionally excluded because default-printer output and modal preview are unsafe for unattended automation.
/// </summary>
[ServiceCategory("workbook", "Workbook")]
[McpTool("workbook", Title = "Workbook Operations", Destructive = true, Category = "structure",
    Description = "TABLE STYLES: create-table-style clones source_style_name without applying it; update-table-style takes a table_style_options object with native elementType names and differential formatting; delete-table-style can remove formatting from existing users. Font name/size, scripts and diagonal borders are unsupported. Built-in styles are read-only. Apply separately with table set-style, pivottable_calc set-layout-options, or slicer set-layout. "
        + "Change workbook document properties, protection and view options, save/copy workbooks, export fixed-format PDF/XPS, update/break external Excel links, and manage native themes and cell styles. CELL STYLES: create-cell-style captures exactly one visible source cell without modifying it; temporarily activates its worksheet and restores the prior view. Hidden source sheets are rejected without changing visibility. update-cell-style changes a custom definition and can affect all existing users throughout the workbook; omitted inclusion flags are preserved. delete-cell-style removes the custom name from existing users; Excel determines retained formatting. Built-in styles are read-only. Apply styles separately with range_format set-style. THEME: apply-theme takes an existing absolute-path .thmx file and changes theme-sensitive formatting throughout the workbook; fixed RGB remains fixed. Saving stays explicit. BREAK-EXTERNAL-LINK HAS NO TOOL-LEVEL UNDO: replaces linked formulas with their current values. SAVE-AS formats: auto, xlsx, xlsm, xlsb, xls; the active session follows the new path. DOCUMENT PROPERTIES: built-in properties can be updated; custom properties can be created, updated, and deleted. Printing and print preview are excluded because default-printer output and modal preview are unsafe for unattended automation.")]
[McpReadOnlyActions("list-table-styles", "get-table-style", "list-cell-styles", "get-cell-style", "get-info",
    "get-theme", "list-document-properties", "get-document-property", "list-external-links", "get-protection", "get-view-options")]
public interface IWorkbookCommands
{
    /// <summary>Lists every native table/Pivot/slicer/timeline style name, built-in/custom status, and availability flags without a cap.</summary>
    [ServiceAction("list-table-styles")]
    TableStyleListResult ListTableStyles(IExcelBatch batch);

    /// <summary>Reads all native elements of the selected table style, including unformatted elements, differential formatting, stripe sizes, and availability. Built-in styles are inspectable but read-only.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="styleName">Native cell or table style name. Discover existing names with list-cell-styles or list-table-styles; create actions require a new name.</param>
    [ServiceAction("get-table-style")]
    TableStyleResult GetTableStyle(IExcelBatch batch, [RequiredParameter] string styleName);

    /// <summary>Creates a custom table style by cloning an existing native definition. It is not applied automatically. Existing names are rejected.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="styleName">New custom style name.</param>
    /// <param name="sourceStyleName">Existing native table/Pivot/slicer/timeline style to clone.</param>
    [ServiceAction("create-table-style")]
    TableStyleResult CreateTableStyle(IExcelBatch batch, [RequiredParameter] string styleName,
        [RequiredParameter] string sourceStyleName);

    /// <summary>Updates custom table/Pivot/slicer/timeline style elements and availability. Omitted settings are preserved; changes can affect existing users throughout the workbook. Built-in styles are read-only. Native failures do not promise rollback.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="styleName">Existing custom style name.</param>
    /// <param name="tableStyleOptions">Typed object: availability flags and elements, each with elementType (native xl name), clear, stripeSize, bold/italic/underline/strikethrough/themeFont, font/fill color/theme/tint, and borders. Only row/column stripes accept stripeSize. No font name/size, script, alignment, number format or diagonals.</param>
    [ServiceAction("update-table-style")]
    TableStyleResult UpdateTableStyle(IExcelBatch batch, [RequiredParameter] string styleName,
        [RequiredParameter] TableStyleOptions tableStyleOptions);

    /// <summary>Deletes a custom table style. Existing tables/Pivots/slicers/timelines can lose its formatting; Excel determines the fallback. Built-in styles cannot be deleted. No tool-level undo.</summary>
    [ServiceAction("delete-table-style")]
    OperationResult DeleteTableStyle(IExcelBatch batch, [RequiredParameter] string styleName);

    /// <summary>Lists every native workbook cell-style name, localized name, and built-in/custom status without a cap. Use get-cell-style for a complete definition.</summary>
    [ServiceAction("list-cell-styles")]
    CellStyleListResult ListCellStyles(IExcelBatch batch);

    /// <summary>Reads the selected cell style's complete native formatting and inclusion flags. Built-in styles are inspectable but not mutable through style lifecycle operations.</summary>
    [ServiceAction("get-cell-style")]
    CellStyleResult GetCellStyle(IExcelBatch batch, [RequiredParameter] string styleName);

    /// <summary>Creates a custom cell style from exactly one visible source cell's stored formatting, without modifying that cell. Temporarily activates its worksheet and restores the prior view; hidden sheets are rejected without changing visibility. Existing names are rejected. Apply it separately with range_format set-style.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="styleName">New custom cell-style name; must not already exist.</param>
    /// <param name="sourceSheetName">Visible source worksheet name.</param>
    /// <param name="sourceCellAddress">Exactly one source cell in the selected workbook.</param>
    [ServiceAction("create-cell-style")]
    CellStyleResult CreateCellStyle(IExcelBatch batch, [RequiredParameter] string styleName,
        [RequiredParameter] string sourceSheetName, [RequiredParameter] string sourceCellAddress);

    /// <summary>Updates a custom cell style. Changes can affect all existing cells using it throughout the workbook. Omitted settings remain unchanged; inclusion flags control which properties applying the style uses. Built-in styles are read-only. Native failures do not promise rollback.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="styleName">Existing custom cell style.</param>
    /// <param name="styleOptions">Typed object: formatOptions (same nested keys as range_format format, except inside borders are range-only), includeFont/includeNumber/includeAlignment/includeBorder/includePatterns/includeProtection, locked, formulaHidden. Omitted inclusion flags are preserved.</param>
    [ServiceAction("update-cell-style")]
    CellStyleResult UpdateCellStyle(IExcelBatch batch, [RequiredParameter] string styleName,
        [RequiredParameter] CellStyleOptions styleOptions);

    /// <summary>Deletes a custom cell style from the workbook. Existing users lose that named style; Excel determines retained cell formatting. Built-in styles cannot be deleted. No tool-level undo.</summary>
    [ServiceAction("delete-cell-style")]
    OperationResult DeleteCellStyle(IExcelBatch batch, [RequiredParameter] string styleName);

    /// <summary>Gets metadata for the active workbook.</summary>
    [ServiceAction("get-info")]
    WorkbookInfoResult GetInfo(IExcelBatch batch);

    /// <summary>Reads all 12 native workbook theme colors and major/minor Latin, East Asian, and complex-script font definitions. Empty script font names remain empty; no fallback font is invented.</summary>
    /// <param name="batch">Excel batch session.</param>
    [ServiceAction("get-theme")]
    WorkbookThemeResult GetTheme(IExcelBatch batch);

    /// <summary>Applies an existing Office .thmx theme through Excel and returns all native theme colors/fonts. This changes theme-sensitive formatting throughout the workbook; fixed RGB colors remain fixed. Saving remains explicit.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="themePath">Absolute path of the existing .thmx file. No theme file is created or copied.</param>
    [ServiceAction("apply-theme")]
    WorkbookThemeResult ApplyTheme(IExcelBatch batch, [RequiredParameter] string themePath);

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

    /// <summary>Permanently breaks one external Excel workbook link, replacing formulas with their current values.
    /// No tool-level undo.</summary>
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
