namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// CLI interface for connection management commands
/// </summary>
public interface IConnectionCommands
{
    int List(string[] args);
    int View(string[] args);
    int Import(string[] args);
    int Export(string[] args);
    int Update(string[] args);
    int Refresh(string[] args);
    int Delete(string[] args);
    int LoadTo(string[] args);
    int GetProperties(string[] args);
    int SetProperties(string[] args);
    int Test(string[] args);
}
