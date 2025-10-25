namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Individual cell operation commands
/// </summary>
public interface ICellCommands
{
    int GetValue(string[] args);
    int SetValue(string[] args);
    int GetFormula(string[] args);
    int SetFormula(string[] args);
    int SetBackgroundColor(string[] args);
    int SetFontColor(string[] args);
    int SetFont(string[] args);
    int SetBorder(string[] args);
    int SetNumberFormat(string[] args);
    int SetAlignment(string[] args);
    int ClearFormatting(string[] args);
}