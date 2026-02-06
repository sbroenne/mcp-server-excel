using Sbroenne.ExcelMcp.ComInterop.Session;
using Sbroenne.ExcelMcp.Core.Attributes;
using Sbroenne.ExcelMcp.Core.Models;

namespace Sbroenne.ExcelMcp.Core.Commands;

/// <summary>
/// Define relationships between Data Model tables for cross-table DAX calculations.
/// Relationships link a foreign key column to a primary key column.
/// </summary>
[ServiceCategory("datamodelrel", "DataModelRel")]
[McpTool("excel_datamodel_rel")]
public interface IDataModelRelCommands
{
    /// <summary>
    /// Lists all table relationships in the model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <returns>Result containing list of relationships</returns>
    [ServiceAction("list-relationships")]
    DataModelRelationshipListResult ListRelationships(IExcelBatch batch);

    /// <summary>
    /// Gets a specific relationship by its table/column identifiers
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <returns>Result containing relationship details</returns>
    [ServiceAction("read-relationship")]
    DataModelRelationshipViewResult ReadRelationship(
        IExcelBatch batch,
        [RequiredParameter] string fromTable,
        [RequiredParameter] string fromColumn,
        [RequiredParameter] string toTable,
        [RequiredParameter] string toColumn);

    /// <summary>
    /// Creates a new relationship between two tables in the Data Model.
    /// Uses Excel COM API: ModelRelationships.Add method (Office 2016+)
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <param name="active">Whether the relationship should be active (default: true)</param>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when tables/columns not found or creation fails</exception>
    [ServiceAction("create-relationship")]
    void CreateRelationship(
        IExcelBatch batch,
        [RequiredParameter] string fromTable,
        [RequiredParameter] string fromColumn,
        [RequiredParameter] string toTable,
        [RequiredParameter] string toColumn,
        bool active = true);

    /// <summary>
    /// Updates an existing relationship's active state in the Data Model.
    /// Uses Excel COM API: ModelRelationship.Active property (Read/Write)
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <param name="active">New active state for the relationship</param>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when relationship not found or update fails</exception>
    [ServiceAction("update-relationship")]
    void UpdateRelationship(
        IExcelBatch batch,
        [RequiredParameter] string fromTable,
        [RequiredParameter] string fromColumn,
        [RequiredParameter] string toTable,
        [RequiredParameter] string toColumn,
        [RequiredParameter] bool active);

    /// <summary>
    /// Deletes a relationship from the Data Model
    /// </summary>
    /// <param name="batch">Excel batch context for accessing workbook</param>
    /// <param name="fromTable">Source table name</param>
    /// <param name="fromColumn">Source column name</param>
    /// <param name="toTable">Target table name</param>
    /// <param name="toColumn">Target column name</param>
    /// <exception cref="ArgumentException">Thrown when parameters are invalid</exception>
    /// <exception cref="InvalidOperationException">Thrown when relationship not found or deletion fails</exception>
    [ServiceAction("delete-relationship")]
    void DeleteRelationship(
        IExcelBatch batch,
        [RequiredParameter] string fromTable,
        [RequiredParameter] string fromColumn,
        [RequiredParameter] string toTable,
        [RequiredParameter] string toColumn);
}
