using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named ranges for formulas/parameters.
/// LIST: returns visible user-defined names; hidden/internal Excel names are omitted before value inspection, and large ranges return metadata without materializing values.
/// CREATE/UPDATE: reference is a cell reference (e.g., 'Sheet1!$A$1').
/// WRITE: value is data to store; invariant numeric and Boolean strings become typed values, while other input remains text.
/// TIP: use range get-values/set-values with the named range as the range address for bulk data read/write.
/// </summary>
[ServiceCategory("namedrange", "NamedRange")]
[MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Partial, false,
    Evidence = "Each named-range action requires explicit native CLI and MCP acceptance.",
    ExcelApiVersion = "Excel Apple Events named item; installed Excel 16.113.1 dictionary.",
    Blocker = "Native named ranges have not completed real CLI and MCP acceptance.")]
[McpTool("namedrange", Title = "Named Range Operations", Destructive = true, Category = "data",
    Description = "Named ranges for formulas/parameters. List returns visible user-defined names; hidden/internal Excel names are omitted before value inspection, and large ranges return metadata without materializing values. Create/update use reference for the cell reference (e.g., Sheet1!$A$1). Write uses value: invariant numeric and Boolean strings become typed values; other input remains text. For bulk data operations, use range with the named range as range_address.")]
public interface INamedRangeCommands
{
    /// <summary>
    /// Lists visible user-defined named ranges in the workbook. Hidden Excel internal names are omitted before value inspection.
    /// Large ranges return metadata without materializing values.
    /// </summary>
    /// <returns>Structured result containing the list of named range information</returns>
    /// <exception cref="InvalidOperationException">If workbook access fails</exception>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "Real CLI/MCP acceptance verifies hidden-name filtering, exact 10000/10001-cell preview bounds and explicit omission for multi-area, constant and ambiguous dynamic names.",
        ExcelApiVersion = "Excel 16.113.1 Apple Events.")]
    [ServiceAction("list")]
    NamedRangeListResult List(IExcelBatch batch);

    /// <summary>
    /// Sets the value of a named range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="name">Name of the named range</param>
    /// <param name="value">Value to set. Invariant numeric and Boolean strings become typed values; other input remains text.</param>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "Real CLI/MCP acceptance verifies scalar types, numeric dates, scoped references and bulk aliases. Shadowed dynamic references are rejected before mutation.",
        ExcelApiVersion = "Excel 16.113.1 Apple Events.")]
    [ServiceAction("write")]
    OperationResult Write(
        IExcelBatch batch,
        [RequiredParameter, FromString("name")] string name,
        [RequiredParameter, FromString("value")] string value);

    /// <summary>
    /// Gets the value of a named range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="name">Name of the named range</param>
    /// <returns>Named range value information</returns>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "Real CLI/MCP acceptance verifies scalar and array values, numeric dates, scoped references, unambiguous dynamic names and persistence.",
        ExcelApiVersion = "Excel 16.113.1 Apple Events.")]
    [ServiceAction("read")]
    NamedRangeValue Read(
        IExcelBatch batch,
        [RequiredParameter, FromString("name")] string name);

    /// <summary>
    /// Updates a named range reference
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="name">Name of the named range</param>
    /// <param name="reference">New cell reference (e.g., Sheet1!$A$1:$B$10)</param>
    /// <exception cref="ArgumentException">If name invalid or too long</exception>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "Real CLI/MCP acceptance verifies normalized references, scalar-to-array updates, quoted worksheet scope and save/reopen.",
        ExcelApiVersion = "Excel 16.113.1 Apple Events.")]
    [ServiceAction("update")]
    OperationResult Update(
        IExcelBatch batch,
        [RequiredParameter, FromString("name")] string name,
        [RequiredParameter, FromString("reference")] string reference);

    /// <summary>
    /// Creates a new named range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="name">Name for the new named range</param>
    /// <param name="reference">Cell reference (e.g., Sheet1!$A$1:$B$10)</param>
    /// <exception cref="ArgumentException">If name invalid or too long</exception>
    /// <exception cref="InvalidOperationException">If named range already exists</exception>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "Real CLI/MCP acceptance verifies workbook-scoped creation and duplicate rejection. Worksheet-scoped creation and local-name collisions are rejected before mutation.",
        ExcelApiVersion = "Excel 16.113.1 typed Apple Events.")]
    [ServiceAction("create")]
    OperationResult Create(
        IExcelBatch batch,
        [RequiredParameter, FromString("name")] string name,
        [RequiredParameter, FromString("reference")] string reference);

    /// <summary>
    /// Deletes a named range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="name">Name of the named range to delete</param>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [MacCapability(MacCapabilityTier.Native, MacImplementationStatus.Implemented, true,
        Evidence = "Real CLI/MCP acceptance verifies workbook and quoted worksheet-scope deletion, missing-name failure and unrelated workbook isolation.",
        ExcelApiVersion = "Excel 16.113.1 typed Apple Events.")]
    [ServiceAction("delete")]
    OperationResult Delete(
        IExcelBatch batch,
        [RequiredParameter, FromString("name")] string name);
}
