using Sbroenne.ExcelMcp.ComInterop;
using Sbroenne.ExcelMcp.Core.DataModel;
using Sbroenne.ExcelMcp.Core.Models;
using Sbroenne.ExcelMcp.Core.Security;
using Sbroenne.ExcelMcp.ComInterop.Session;

#pragma warning disable CS1998 // Async method lacks 'await' operators - intentional for COM synchronous operations

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model management commands - Core data layer (no console output)
/// Provides read-only access to Excel Data Model (PowerPivot) objects
/// Split into partial classes: Base (constructor), Read (List/View/Export), Write (Delete/Create/Update), Refresh
/// </summary>
public partial class DataModelCommands : IDataModelCommands
{
    /// <summary>
    /// Constructor for DataModelCommands
    /// </summary>
    public DataModelCommands()
    {
        // No dependencies currently needed
    }
}
