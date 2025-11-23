using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Named range/parameter management commands
/// </summary>
public interface INamedRangeCommands
{
    /// <summary>
    /// Lists all named ranges in the workbook
    /// </summary>
    /// <returns>List of named range information</returns>
    /// <exception cref="InvalidOperationException">If workbook access fails</exception>
    List<NamedRangeInfo> List(IExcelBatch batch);

    /// <summary>
    /// Sets the value of a named range
    /// </summary>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    void Write(IExcelBatch batch, string paramName, string value);

    /// <summary>
    /// Gets the value of a named range
    /// </summary>
    /// <returns>Named range value information</returns>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    NamedRangeValue Read(IExcelBatch batch, string paramName);

    /// <summary>
    /// Updates a named range reference
    /// </summary>
    /// <exception cref="ArgumentException">If parameter name invalid or too long</exception>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    void Update(IExcelBatch batch, string paramName, string reference);

    /// <summary>
    /// Creates a new named range
    /// </summary>
    /// <exception cref="ArgumentException">If parameter name invalid or too long</exception>
    /// <exception cref="InvalidOperationException">If named range already exists</exception>
    void Create(IExcelBatch batch, string paramName, string reference);

    /// <summary>
    /// Deletes a named range
    /// </summary>
    /// <exception cref="InvalidOperationException">If named range not found</exception>
    void Delete(IExcelBatch batch, string paramName);
}

