using Sbroenne.ExcelMcp.ComInterop.Session;
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
    /// Disable VBA project access trust in Excel
    /// </summary>
    VbaTrustResult DisableVbaTrust();

    /// <summary>
    /// Check current VBA trust status
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    Task<VbaTrustResult> CheckVbaTrustAsync(IExcelBatch batch);
}
