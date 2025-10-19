using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Setup and configuration commands for ExcelCLI
/// </summary>
public interface ISetupCommands
{
    /// <summary>
    /// Enable VBA project access trust in Excel
    /// </summary>
    VbaTrustResult EnableVbaTrust();
    
    /// <summary>
    /// Check current VBA trust status
    /// </summary>
    /// <param name="testFilePath">Path to Excel file to test VBA access</param>
    VbaTrustResult CheckVbaTrust(string testFilePath);
}
