namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Named range/parameter management commands
/// </summary>
public interface INamedRangeCommands
{
    int List(string[] args);
    int SetValue(string[] args);
    int GetValue(string[] args);
    int Update(string[] args);
    int Create(string[] args);
    int Delete(string[] args);
}
