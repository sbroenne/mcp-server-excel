namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Individual cell operation commands
/// </summary>
public interface ICellCommands
{
    /// <summary>
    /// Gets the value of a specific cell
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheet, cellAddress]</param>
    /// <returns>0 on success, 1 on error</returns>
    int GetValue(string[] args);
    
    /// <summary>
    /// Sets the value of a specific cell
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheet, cellAddress, value]</param>
    /// <returns>0 on success, 1 on error</returns>
    int SetValue(string[] args);
    
    /// <summary>
    /// Gets the formula of a specific cell
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheet, cellAddress]</param>
    /// <returns>0 on success, 1 on error</returns>
    int GetFormula(string[] args);
    
    /// <summary>
    /// Sets the formula of a specific cell
    /// </summary>
    /// <param name="args">Command arguments: [file.xlsx, sheet, cellAddress, formula]</param>
    /// <returns>0 on success, 1 on error</returns>
    int SetFormula(string[] args);
}
