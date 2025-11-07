namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Power Query management commands
/// </summary>
public interface IPowerQueryCommands
{
    int List(string[] args);
    int View(string[] args);
    Task<int> Export(string[] args);
    int Refresh(string[] args);
    int Delete(string[] args);
    int Sources(string[] args);
    int Test(string[] args);
    int Peek(string[] args);
    int Eval(string[] args);
}
