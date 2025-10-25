using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Interface for hyperlink-related commands in Excel workbooks.
/// </summary>
public interface IHyperlinkCommands
{
    /// <summary>
    /// Adds a hyperlink to a cell or range in a worksheet.
    /// </summary>
    /// <param name="excelPath">Path to the Excel file</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell address (e.g., "A1") or range</param>
    /// <param name="url">URL or file path for the hyperlink</param>
    /// <param name="displayText">Optional display text (defaults to cell value or URL)</param>
    /// <param name="tooltip">Optional tooltip/screen tip text</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult AddHyperlink(string excelPath, string sheetName, string cellAddress, string url, string? displayText = null, string? tooltip = null);

    /// <summary>
    /// Removes a hyperlink from a cell or range in a worksheet.
    /// </summary>
    /// <param name="excelPath">Path to the Excel file</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell address (e.g., "A1") or range</param>
    /// <returns>Result indicating success or failure</returns>
    OperationResult RemoveHyperlink(string excelPath, string sheetName, string cellAddress);

    /// <summary>
    /// Lists all hyperlinks in a worksheet.
    /// </summary>
    /// <param name="excelPath">Path to the Excel file</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <returns>Result containing list of hyperlinks</returns>
    HyperlinkListResult ListHyperlinks(string excelPath, string sheetName);

    /// <summary>
    /// Gets hyperlink information for a specific cell.
    /// </summary>
    /// <param name="excelPath">Path to the Excel file</param>
    /// <param name="sheetName">Name of the worksheet</param>
    /// <param name="cellAddress">Cell address (e.g., "A1")</param>
    /// <returns>Result containing hyperlink details</returns>
    HyperlinkInfoResult GetHyperlink(string excelPath, string sheetName, string cellAddress);
}
