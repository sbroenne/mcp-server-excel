using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named ranges for formulas/parameters.
/// CREATE/UPDATE: value is cell reference (e.g., 'Sheet1!$A$1').
/// WRITE: value is data to store.
/// TIP: range(rangeAddress=namedRangeName) for bulk data read/write.
/// </summary>
[ServiceCategory("namedrange", "NamedRange")]
[McpTool("namedrange", Title = "Named Range Operations", Destructive = true, Category = "data",
    Description = "Named ranges for formulas/parameters. CREATE/UPDATE: value is cell reference (e.g., Sheet1!$A$1). WRITE: value is data to store in the named range. TIP: Use range(rangeAddress=namedRangeName) for bulk data operations.")]
public interface INamedRangeCommands
{
    /// <summary>
    /// Lists all named ranges in the workbook
    /// </summary>
    /// <returns>List of named range information</returns>
    /// <exception cref="InvalidOperationException">If workbook access fails</exception>
    [ServiceAction("list")]
    List<NamedRangeInfo> List(IExcelBatch batch);

    /// <summary>
    /// Sets the value of a named range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="name">Name of the named range</param>
    /// <param name="value">Value to set</param>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [ServiceAction("write")]
    void Write(
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
    [ServiceAction("update")]
    void Update(
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
    [ServiceAction("create")]
    void Create(
        IExcelBatch batch,
        [RequiredParameter, FromString("name")] string name,
        [RequiredParameter, FromString("reference")] string reference);

    /// <summary>
    /// Deletes a named range
    /// </summary>
    /// <param name="batch">Excel batch session</param>
    /// <param name="name">Name of the named range to delete</param>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [ServiceAction("delete")]
    void Delete(
        IExcelBatch batch,
        [RequiredParameter, FromString("name")] string name);
}



