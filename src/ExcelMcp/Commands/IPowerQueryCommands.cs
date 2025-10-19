namespace ExcelMcp.Commands;

/// <summary>
/// Power Query management commands
/// </summary>
public interface IPowerQueryCommands
{
    int List(string[] args);
    int View(string[] args);
    Task<int> Update(string[] args);
    Task<int> Export(string[] args);
    Task<int> Import(string[] args);
    int Refresh(string[] args);
    int Errors(string[] args);
    int LoadTo(string[] args);
    int Delete(string[] args);
}
