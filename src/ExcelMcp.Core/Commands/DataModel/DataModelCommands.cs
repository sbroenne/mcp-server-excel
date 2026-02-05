
namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Data Model management commands - Core data layer (no console output)
/// Provides access to Excel Data Model (PowerPivot) objects: tables, measures, relationships.
/// Split into partial classes: Base (constructor), Read (List/View/Export), Write (Delete/Create/Update), Refresh
/// Implements both IDataModelCommands (tables/measures) and IDataModelRelCommands (relationships).
/// </summary>
public partial class DataModelCommands : IDataModelCommands, IDataModelRelCommands
{
    /// <summary>
    /// Constructor for DataModelCommands
    /// </summary>
    public DataModelCommands()
    {
        // No dependencies currently needed
    }
}


