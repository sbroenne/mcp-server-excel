namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Worksheet management commands
/// </summary>
public interface ISheetCommands
{
    int List(string[] args);
    int Read(string[] args);
    Task<int> Write(string[] args);
    int Copy(string[] args);
    int Delete(string[] args);
    int Create(string[] args);
    int Rename(string[] args);
    int Clear(string[] args);
    int Append(string[] args);
}
