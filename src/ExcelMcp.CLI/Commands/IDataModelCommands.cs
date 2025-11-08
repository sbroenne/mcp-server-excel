namespace Sbroenne.ExcelMcp.CLI.Commands;

/// <summary>
/// Interface for Data Model CLI commands
/// </summary>
public interface IDataModelCommands
{
    /// <summary>
    /// Lists all tables in the Data Model
    /// Usage: dm-list-tables &lt;file.xlsx&gt;
    /// </summary>
    int ListTables(string[] args);

    /// <summary>
    /// Lists all DAX measures in the Data Model
    /// Usage: dm-list-measures &lt;file.xlsx&gt;
    /// </summary>
    int ListMeasures(string[] args);

    /// <summary>
    /// Views a specific DAX measure formula
    /// Usage: dm-view-measure &lt;file.xlsx&gt; &lt;measure-name&gt;
    /// </summary>
    int ViewMeasure(string[] args);

    /// <summary>
    /// Exports a DAX measure to a file
    /// Usage: dm-export-measure &lt;file.xlsx&gt; &lt;measure-name&gt; &lt;output.dax&gt;
    /// </summary>
    int ExportMeasure(string[] args);

    /// <summary>
    /// Lists all relationships in the Data Model
    /// Usage: dm-list-relationships &lt;file.xlsx&gt;
    /// </summary>
    int ListRelationships(string[] args);

    /// <summary>
    /// Refreshes the Data Model
    /// Usage: dm-refresh &lt;file.xlsx&gt;
    /// </summary>
    int Refresh(string[] args);

    /// <summary>
    /// Deletes a DAX measure from the Data Model
    /// Usage: dm-delete-measure &lt;file.xlsx&gt; &lt;measure-name&gt;
    /// </summary>
    int DeleteMeasure(string[] args);

    /// <summary>
    /// Deletes a relationship from the Data Model
    /// Usage: dm-delete-relationship &lt;file.xlsx&gt; &lt;from-table&gt; &lt;from-column&gt; &lt;to-table&gt; &lt;to-column&gt;
    /// </summary>
    int DeleteRelationship(string[] args);

    // Phase 2: Discovery operations

    /// <summary>
    /// Lists all columns in a Data Model table
    /// Usage: dm-list-columns &lt;file.xlsx&gt; &lt;table-name&gt;
    /// </summary>
    int ListColumns(string[] args);

    /// <summary>
    /// Views detailed information about a Data Model table
    /// Usage: dm-view-table &lt;file.xlsx&gt; &lt;table-name&gt;
    /// </summary>
    int ViewTable(string[] args);

    /// <summary>
    /// Gets Data Model overview (table/measure/relationship counts)
    /// Usage: dm-get-model-info &lt;file.xlsx&gt;
    /// </summary>
    int GetModelInfo(string[] args);

    // Phase 2: CREATE/UPDATE operations

    /// <summary>
    /// Creates a new DAX measure in the Data Model
    /// Usage: dm-create-measure &lt;file.xlsx&gt; &lt;table-name&gt; &lt;measure-name&gt; &lt;dax-formula&gt; [format-type] [description]
    /// </summary>
    int CreateMeasure(string[] args);

    /// <summary>
    /// Updates an existing DAX measure
    /// Usage: dm-update-measure &lt;file.xlsx&gt; &lt;measure-name&gt; [dax-formula] [format-type] [description]
    /// </summary>
    int UpdateMeasure(string[] args);

    /// <summary>
    /// Creates a relationship between two Data Model tables
    /// Usage: dm-create-relationship &lt;file.xlsx&gt; &lt;from-table&gt; &lt;from-column&gt; &lt;to-table&gt; &lt;to-column&gt; [active:true|false]
    /// </summary>
    int CreateRelationship(string[] args);

    /// <summary>
    /// Updates a relationship's active status
    /// Usage: dm-update-relationship &lt;file.xlsx&gt; &lt;from-table&gt; &lt;from-column&gt; &lt;to-table&gt; &lt;to-column&gt; &lt;active:true|false&gt;
    /// </summary>
    int UpdateRelationship(string[] args);
}

