namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Worksheet lifecycle management commands
/// Data operations (read, write, clear, append) moved to RangeCommands in Phase 1A.
/// </summary>
public interface ISheetCommands
{
    int List(string[] args);
    int Create(string[] args);
    int Rename(string[] args);
    int Copy(string[] args);
    int Delete(string[] args);
}
