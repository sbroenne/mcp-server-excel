namespace ExcelMcp.Commands;

/// <summary>
/// Named range/parameter management commands
/// </summary>
public interface IParameterCommands
{
    int List(string[] args);
    int Set(string[] args);
    int Get(string[] args);
    int Create(string[] args);
    int Delete(string[] args);
}
