namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Setup and configuration commands for ExcelCLI
/// </summary>
public interface ISetupCommands
{
    /// <summary>
    /// Enable VBA project access trust in Excel
    /// </summary>
    int EnableVbaTrust(string[] args);
    
    /// <summary>
    /// Check current VBA trust status
    /// </summary>
    int CheckVbaTrust(string[] args);
}
