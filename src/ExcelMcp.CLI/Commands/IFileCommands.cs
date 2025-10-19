namespace ExcelMcp.Commands;

/// <summary>
/// File management commands for Excel workbooks
/// </summary>
public interface IFileCommands
{
    int CreateEmpty(string[] args);
}