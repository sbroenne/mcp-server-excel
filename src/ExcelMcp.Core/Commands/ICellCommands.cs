namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Individual cell operation commands
/// </summary>
public interface ICellCommands
{
    int GetValue(string[] args);
    int SetValue(string[] args);
    int GetFormula(string[] args);
    int SetFormula(string[] args);
}
