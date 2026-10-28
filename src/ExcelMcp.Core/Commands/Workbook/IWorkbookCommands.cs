using Sbroenne.ExcelMcp.ComInterop;
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
[ServiceCategory("Workbook")]
[McpTool("workbook", Title = "Workbook Operations", Destructive = true, Category = "structure",
    Description = "Save, copy, export, refresh, and protect workbooks; manage document properties, view options, external links, themes, table styles, and cell styles. " +
        "REFRESH ALL: refreshes every data connection and PivotTable, waits for completion, and reports refreshed and failed items. " +
        "SAVE-AS formats: auto, xlsx, xlsm, xlsb, xls; the session follows the new path. save-as and save-copy-as write files even with AutoSave off; other actions leave edits unsaved unless AutoSave is on, and enabling AutoSave can save pending edits immediately. Printing is not supported. " +
        "BREAK-EXTERNAL-LINK HAS NO TOOL-LEVEL UNDO: replaces linked formulas with their current values. " +
        "DOCUMENT INSPECTOR: inspect-document runs every available native module without removing anything or saving, but can mark a saved workbook unsaved (reports savedBefore/savedAfter). fix-document-inspector requires inspector_index and the exact inspector_name from inspect-document and removes that whole category, possibly including worksheets and data, with no tool-level undo or automatic save. Not a full privacy audit: comments/revisions and document properties are excluded. " +
        "TABLE STYLES: create-table-style clones source_style_name without applying it; update-table-style takes table_style_options with native elementType names; Font name/size, scripts, and diagonal borders are unsupported. Apply with table set-style, pivottable_calc set-layout-options, or slicer set-layout. " +
        "CELL STYLES: create-cell-style copies one visible source cell; update-cell-style can change all existing users and keeps omitted inclusion flags. Apply with range_format set-style. Deleting a table or cell style removes it from existing users. Built-in styles are read-only. " +
        "THEME: apply-theme takes an absolute path to an existing .thmx file and changes theme-sensitive formatting throughout the workbook; fixed RGB colors stay.")]
[McpReadOnlyActions("inspect", "list-table-styles", "get-table-style", "list-cell-styles", "get-cell-style", "get-info",
    "get-theme", "list-document-properties", "get-document-property", "list-external-links", "get-protection", "get-view-options")]
public interface IWorkbookCommands
{
    /// <summary>
    /// Refreshes all Power Query loads and other workbook data connections, then every
    /// PivotTable. Refreshes synchronously, continues after individual failures, and
    /// reports the refreshed and failed objects without rolling back successful refreshes.
    /// </summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="timeout">Maximum time to wait; omitted or zero uses the data-operation default.</param>
    /// <param name="progress">Optional progress reporter.</param>
    [ServiceAction("refresh-all")]
    WorkbookRefreshAllResult RefreshAll(
        IExcelBatch batch,
        TimeSpan timeout = default,
        IProgress<ProgressInfo>? progress = null);

    /// <summary>
    /// Runs every available native Document Inspector module without requesting edits or saving.
    /// Returns each module's index, localized name, native status and result text.
    /// Reports savedBefore/savedAfter because native modules can mark a saved workbook unsaved even without a fix.
    /// Not a full privacy audit: comments/revisions and document properties are not exposed by this collection.
    /// Module availability depends on the installed Excel version and custom inspectors.
    /// Findings are successful inspections; native errors or no available modules are failures.
    /// Distinct from inspect, which returns a bounded workbook content overview.
    /// </summary>
    [ServiceAction("inspect-document")]
    DocumentInspectionResult InspectDocument(IExcelBatch batch);

    /// <summary>
    /// Runs the native fix for exactly one selected Document Inspector module and reinspects it.
    /// May remove every item in the selected category, including hidden worksheets and their data.
    /// No tool-level undo or automatic save; existing Excel AutoSave may persist edits.
    /// Rejects read-only or structure/window-protected workbooks before mutation.
    /// Failure may leave partial changes; inspect the workbook before deciding what to do next.
    /// </summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="inspectorIndex">One-based inspector index returned by inspect-document; not a fixed category identifier.</param>
    /// <param name="inspectorName">Exact localized native name returned by inspect-document, including any ampersands. Must still match the selected index.</param>
    [ServiceAction("fix-document-inspector")]
    DocumentInspectorFixResult FixDocumentInspector(IExcelBatch batch,
        [RequiredParameter] int inspectorIndex, [RequiredParameter] string inspectorName);

    /// <summary>
    /// Returns bounded workbook structure and optionally a bounded preview from one selected worksheet.
    /// Hidden worksheets are included with their visibility. Counts and omitted-item totals describe the selected scope.
    /// Preview limits are applied before reading cells; string text is capped per cell and across the whole preview.
    /// Does not edit, save, refresh, recalculate, run macros, or change the active view.
    /// </summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="sheetName">Optional worksheet filter for sheet/table metadata; required when includePreview is true.</param>
    /// <param name="includeSheets">Include worksheet names, visibility, and used-range geometry.</param>
    /// <param name="includeTables">Include table names, worksheets, and ranges.</param>
    /// <param name="includeDefinedNames">Include visible user-defined workbook names and their RefersTo expressions.</param>
    /// <param name="includePreview">Include a bounded value/formula preview. Requires sheetName.</param>
    /// <param name="rangeAddress">Optional preview range on sheetName. If omitted, previews the worksheet's UsedRange from its top-left cell.</param>
    /// <param name="maxItems">Maximum returned items per metadata section, 1-100.</param>
    /// <param name="maxPreviewRows">Maximum preview rows, 1-10.</param>
    /// <param name="maxPreviewColumns">Maximum preview columns, 1-10.</param>
    /// <param name="maxCellCharacters">Maximum text characters per value/formula cell, 1-200.</param>
    /// <param name="maxPreviewCharacters">Maximum combined text characters in the preview values and formulas, 1-8192.</param>
    [ServiceAction("inspect")]
    WorkbookOverviewResult Inspect(
        IExcelBatch batch,
        string? sheetName = null,
        bool includeSheets = true,
        bool includeTables = true,
        bool includeDefinedNames = true,
        bool includePreview = false,
        string? rangeAddress = null,
        int maxItems = 50,
        int maxPreviewRows = 6,
        int maxPreviewColumns = 6,
        int maxCellCharacters = 160,
        int maxPreviewCharacters = 4096);

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

    /// <summary>Creates a custom cell style from exactly one visible source cell's stored formatting, without modifying that cell. Temporarily activates its worksheet and restores the prior view; hidden sheets are rejected without changing visibility. Existing names are rejected. Apply it separately with range_format set-style (CLI: rangeformat set-style).</summary>
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

    /// <summary>Gets metadata for the active workbook, including saved, readOnly, and the live autoSaveOn status. AutoSave can persist edits before an explicit save when enabled.</summary>
    [ServiceAction("get-info")]
    WorkbookInfoResult GetInfo(IExcelBatch batch);

    /// <summary>Changes AutoSave only for the selected workbook and returns its confirmed status. Enabling AutoSave can immediately save pending edits. Unsupported settings and native failures are errors; no rollback is promised.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="enabled">Whether Excel AutoSave should be enabled.</param>
    [ServiceAction("set-auto-save")]
    WorkbookInfoResult SetAutoSave(IExcelBatch batch, [RequiredParameter] bool enabled);

    /// <summary>Reads all 12 native workbook theme colors and major/minor Latin, East Asian, and complex-script font definitions. Empty script font names remain empty; no fallback font is invented.</summary>
    /// <param name="batch">Excel batch session.</param>
    [ServiceAction("get-theme")]
    WorkbookThemeResult GetTheme(IExcelBatch batch);

    /// <summary>Applies an existing Office .thmx theme through Excel and returns all native theme colors/fonts. This changes theme-sensitive formatting throughout the workbook; fixed RGB colors remain fixed. Saving remains explicit.</summary>
    /// <param name="batch">Excel batch session.</param>
    /// <param name="themePath">Absolute path of the existing .thmx file. No theme file is created or copied.</param>
    [ServiceAction("apply-theme")]
    WorkbookThemeResult ApplyTheme(IExcelBatch batch, [RequiredParameter] string themePath);

    /// <summary>Lists built-in and/or custom workbook document properties. Each property reports valueAvailable and valueError. An unreadable built-in value can be undefined or its getter can have failed; valueAvailable=false does not prove absence. Inspect the native error rather than interpreting a null value as empty metadata.</summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="includeBuiltIn">Include built-in document properties</param>
    /// <param name="includeCustom">Include custom document properties</param>
    [ServiceAction("list-document-properties")]
    DocumentPropertyListResult ListDocumentProperties(
        IExcelBatch batch,
        bool includeBuiltIn = true,
        bool includeCustom = true);

    /// <summary>Gets one built-in or custom workbook document property, including valueAvailable and valueError. An unreadable value is explicitly reported, not treated as confirmed absence; unexpected native errors still fail the operation.</summary>
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
