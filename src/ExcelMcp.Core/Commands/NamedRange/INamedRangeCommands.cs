using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands
/// </summary>
[ServiceCategory("namedrange", "NamedRange")]
[McpTool("excel_namedrange")]
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
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [ServiceAction("write")]
    void Write(
        IExcelBatch batch,
        [RequiredParameter, FromString("paramName")] string paramName,
        [RequiredParameter, FromString("value")] string value);

    /// <summary>
    /// Gets the value of a named range
    /// </summary>
    /// <returns>Named range value information</returns>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [ServiceAction("read")]
    NamedRangeValue Read(
        IExcelBatch batch,
        [RequiredParameter, FromString("paramName")] string paramName);

    /// <summary>
    /// Updates a named range reference
    /// </summary>
    /// <exception cref="ArgumentException">If parameter name invalid or too long</exception>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [ServiceAction("update")]
    void Update(
        IExcelBatch batch,
        [RequiredParameter, FromString("paramName")] string paramName,
        [RequiredParameter, FromString("reference")] string reference);

    /// <summary>
    /// Creates a new named range
    /// </summary>
    /// <exception cref="ArgumentException">If parameter name invalid or too long</exception>
    /// <exception cref="InvalidOperationException">If named range already exists</exception>
    [ServiceAction("create")]
    void Create(
        IExcelBatch batch,
        [RequiredParameter, FromString("paramName")] string paramName,
        [RequiredParameter, FromString("reference")] string reference);

    /// <summary>
    /// Deletes a named range
    /// </summary>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    [ServiceAction("delete")]
    void Delete(
        IExcelBatch batch,
        [RequiredParameter, FromString("paramName")] string paramName);
}



