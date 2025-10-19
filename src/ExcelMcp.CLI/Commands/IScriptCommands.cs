namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// VBA script management commands
/// </summary>
public interface IScriptCommands
{
    int List(string[] args);
    int Export(string[] args);
    Task<int> Import(string[] args);
    Task<int> Update(string[] args);
    int Run(string[] args);
    int Delete(string[] args);
}